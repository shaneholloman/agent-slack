import type { SlackApiClient, SlackAuth } from "./client.ts";
import { downloadSlackFile } from "./files.ts";
import { htmlToMarkdown } from "./html-to-md.ts";
import { ensureDownloadsDir } from "../lib/tmp-paths.ts";
import { getString, isRecord } from "../lib/object-type-guards.ts";
import { readFile } from "node:fs/promises";
import { getUserAgent } from "../lib/version.ts";

export type SlackCanvasRef = {
  workspace_url: string;
  canvas_id: string; // looks like a file id, e.g. F080JDE025R
  raw: string;
};

export const CANVAS_EDIT_OPERATIONS = [
  "insert_after",
  "insert_before",
  "insert_at_start",
  "insert_at_end",
  "replace",
  "delete",
  "rename",
] as const;

export const MAX_CANVAS_EDIT_MARKDOWN_CHARS = 1_048_576;

export type SlackCanvasEditOperation = (typeof CANVAS_EDIT_OPERATIONS)[number];

export type SlackCanvasEditChange = {
  operation: SlackCanvasEditOperation;
  section_id?: string;
  document_content?: { type: "markdown"; markdown: string };
  title_content?: { type: "markdown"; markdown: string };
};

export const BROWSER_AUTH_CANVAS_CHANNEL_ERROR =
  "Adding a canvas as a channel tab requires a standard Slack token; imported browser credentials can create standalone canvases only";

export const BROWSER_AUTH_CANVAS_EDIT_ERROR =
  "Editing a canvas requires a standard Slack token with the canvases:write scope; imported browser credentials cannot call canvases.edit";

export async function createCanvasFromMarkdown(
  client: SlackApiClient,
  input: {
    auth: SlackAuth;
    markdown: string;
    title?: string;
    channelId?: string;
  },
): Promise<{ canvas: { id: string; title?: string; channel_id?: string } }> {
  if (!input.markdown.trim()) {
    throw new Error("Canvas Markdown is empty");
  }

  const title = input.title?.trim() || undefined;
  let response: Record<string, unknown>;
  let canvasId: string | undefined;

  if (input.auth.auth_type === "browser") {
    if (input.channelId) {
      throw new Error(BROWSER_AUTH_CANVAS_CHANNEL_ERROR);
    }
    response = await client.apiMultipart("files.createCanvas", {
      title: title ?? "Untitled",
      markdown: input.markdown,
      loosenValidation: true,
    });
    canvasId = getString(response.file_id);
  } else {
    response = await client.api("canvases.create", {
      title,
      document_content: {
        type: "markdown",
        markdown: input.markdown,
      },
      channel_id: input.channelId,
    });
    canvasId = getString(response.canvas_id);
  }

  if (!canvasId) {
    throw new Error("Slack returned no canvas id");
  }

  return {
    canvas: {
      id: canvasId,
      title: title ?? (input.auth.auth_type === "browser" ? "Untitled" : undefined),
      channel_id: input.channelId,
    },
  };
}

function isCanvasEditOperation(value: string): value is SlackCanvasEditOperation {
  return (CANVAS_EDIT_OPERATIONS as readonly string[]).includes(value);
}

/**
 * Apply one documented Slack Canvas edit operation.
 *
 * Slack currently accepts exactly one change per canvases.edit call. Keeping
 * the operation validation here makes the library safe for callers other than
 * the CLI and prevents malformed requests from reaching Slack.
 */
export async function editCanvas(
  client: SlackApiClient,
  input: {
    canvasId: string;
    operation: string;
    markdown?: string;
    title?: string;
    sectionId?: string;
  },
): Promise<{ ok: true; canvas: { id: string; operation: SlackCanvasEditOperation } }> {
  const canvasId = input.canvasId.trim();
  if (!canvasId) {
    throw new Error("Canvas id is required");
  }

  const operation = input.operation.trim();
  if (!isCanvasEditOperation(operation)) {
    throw new Error(
      `Unsupported canvas edit operation "${input.operation}". Expected one of: ${CANVAS_EDIT_OPERATIONS.join(", ")}`,
    );
  }

  const sectionId = input.sectionId?.trim() || undefined;
  const { markdown } = input;
  const title = input.title?.trim() || undefined;
  const contentOperations = new Set<SlackCanvasEditOperation>([
    "insert_after",
    "insert_before",
    "insert_at_start",
    "insert_at_end",
    "replace",
  ]);

  if (contentOperations.has(operation)) {
    if (markdown === undefined || !markdown.trim()) {
      throw new Error(`Canvas edit operation "${operation}" requires non-empty Markdown content`);
    }
    if (markdown.length > MAX_CANVAS_EDIT_MARKDOWN_CHARS) {
      throw new Error(
        `Canvas edit Markdown exceeds Slack's ${MAX_CANVAS_EDIT_MARKDOWN_CHARS.toLocaleString("en-US")} character limit`,
      );
    }
    if (title !== undefined) {
      throw new Error(`--title is only valid with the rename operation`);
    }
    if ((operation === "insert_after" || operation === "insert_before") && !sectionId) {
      throw new Error(`Canvas edit operation "${operation}" requires --section-id`);
    }
    if (
      (operation === "insert_at_start" || operation === "insert_at_end") &&
      sectionId !== undefined
    ) {
      throw new Error(`Canvas edit operation "${operation}" cannot use --section-id`);
    }
  } else if (operation === "delete") {
    if (!sectionId) {
      throw new Error('Canvas edit operation "delete" requires --section-id');
    }
    if (markdown !== undefined) {
      throw new Error('Canvas edit operation "delete" does not accept Markdown content');
    }
    if (title !== undefined) {
      throw new Error(`--title is only valid with the rename operation`);
    }
  } else {
    if (!title) {
      throw new Error('Canvas edit operation "rename" requires a non-empty --title');
    }
    if (title.length > MAX_CANVAS_EDIT_MARKDOWN_CHARS) {
      throw new Error(
        `Canvas edit title exceeds Slack's ${MAX_CANVAS_EDIT_MARKDOWN_CHARS.toLocaleString("en-US")} character limit`,
      );
    }
    if (markdown !== undefined) {
      throw new Error('Canvas edit operation "rename" does not accept Markdown content');
    }
    if (sectionId !== undefined) {
      throw new Error('Canvas edit operation "rename" cannot use --section-id');
    }
  }

  const change: SlackCanvasEditChange = { operation };
  if (sectionId !== undefined) {
    change.section_id = sectionId;
  }
  if (contentOperations.has(operation)) {
    change.document_content = { type: "markdown", markdown: markdown! };
  } else if (operation === "rename") {
    change.title_content = { type: "markdown", markdown: title! };
  }

  await client.api("canvases.edit", {
    canvas_id: canvasId,
    changes: [change],
  });

  return { ok: true, canvas: { id: canvasId, operation } };
}

export function parseSlackCanvasUrl(input: string): SlackCanvasRef {
  let url: URL;
  try {
    url = new URL(input);
  } catch {
    throw new Error(`Invalid URL: ${input}`);
  }

  if (!/\.slack\.com$/i.test(url.hostname)) {
    throw new Error(`Not a Slack workspace URL: ${url.hostname}`);
  }

  // Common form: /docs/<team_id>/<canvas_id>
  // Example seen in Slack docs: https://workspace.slack.com/docs/T.../F...
  const parts = url.pathname.split("/").filter(Boolean);
  if (parts[0] !== "docs") {
    throw new Error(`Unsupported Slack canvas URL path: ${url.pathname}`);
  }

  const canvas_id = parts.find((p) => /^F[A-Z0-9]{8,}$/.test(p));
  if (!canvas_id) {
    throw new Error(`Could not find canvas id in: ${url.pathname}`);
  }

  const workspace_url = `${url.protocol}//${url.host}`;
  return { workspace_url, canvas_id, raw: input };
}

export async function fetchCanvasMarkdown(
  client: SlackApiClient,
  input: {
    auth: SlackAuth;
    workspaceUrl: string;
    canvasId: string;
    options?: { maxChars?: number; downloadHtml?: boolean };
  },
): Promise<{ canvas: { id: string; title?: string; markdown: string } }> {
  const info = await client.api("files.info", { file: input.canvasId });
  const file = isRecord(info.file) ? info.file : null;
  if (!file) {
    throw new Error("Canvas not found (files.info returned no file)");
  }

  const title = (getString(file.title) || getString(file.name) || "").trim() || undefined;
  const downloadUrl = getString(file.url_private_download) ?? getString(file.url_private);
  if (!downloadUrl) {
    throw new Error("Canvas has no download URL");
  }

  let html = "";
  if (input.options?.downloadHtml ?? true) {
    const htmlPath = await downloadSlackFile({
      auth: input.auth,
      url: downloadUrl,
      // keep canvases with other downloads (agent-friendly temp dir)
      // filename uses canvasId (unique)
      // Note: canvases download as HTML via Slack file endpoints
      // so allowHtml must be true.
      destDir: await ensureDownloadsDir(),
      preferredName: `${input.canvasId}.html`,
      options: { allowHtml: true },
    });
    html = await readFile(htmlPath, "utf8");
  } else {
    const headers: Record<string, string> = {};
    if (input.auth.auth_type === "standard") {
      headers.Authorization = `Bearer ${input.auth.token}`;
    } else {
      headers.Authorization = `Bearer ${input.auth.xoxc_token}`;
      headers.Cookie = `d=${encodeURIComponent(input.auth.xoxd_cookie)}`;
      headers.Referer = "https://app.slack.com/";
      headers["User-Agent"] = getUserAgent();
    }
    const resp = await fetch(downloadUrl, { headers });
    if (!resp.ok) {
      throw new Error(`Failed to download canvas HTML (${resp.status})`);
    }
    html = await resp.text();
  }

  const markdownRaw = htmlToMarkdown(html).trim();
  const maxChars = input.options?.maxChars ?? 20000;
  const markdown =
    maxChars >= 0 && markdownRaw.length > maxChars
      ? `${markdownRaw.slice(0, maxChars)}\n…`
      : markdownRaw;

  return {
    canvas: {
      id: input.canvasId,
      title,
      markdown,
    },
  };
}
