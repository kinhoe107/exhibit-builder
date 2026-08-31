import PostalMime from "postal-mime";

const EMAIL_LIMITS = {
  rawBytes: 32 * 1024 * 1024,
  bodyBytes: 16 * 1024 * 1024,
  attachments: 500,
  milliseconds: 30_000,
} as const;

export type ParsedEmailAttachment = {
  name: string;
  mimeType: string;
  disposition: string;
  content: Uint8Array;
};

export type ParsedBundleEmail = {
  headers: Map<string, string>;
  body: string;
  attachments: ParsedEmailAttachment[];
};

function htmlToText(html: string) {
  return html
    .replace(/<\s*(script|style|iframe|object)[^>]*>[\s\S]*?<\s*\/\s*\1\s*>/gi, "")
    .replace(/<\s*br\s*\/?>/gi, "\n")
    .replace(/<\s*\/\s*(p|div|li|tr|h[1-6])\s*>/gi, "\n")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/\n{3,}/g, "\n\n")
    .replace(/[ \t]+/g, " ")
    .trim();
}

/** RFC822/MIME parsing with bounded nesting/header sizes; no HTML is executed. */
export async function parseBundleEmail(raw: string): Promise<ParsedBundleEmail> {
  if (new TextEncoder().encode(raw).byteLength > EMAIL_LIMITS.rawBytes) throw new Error("Email exceeds the 32 MB parsing safety limit.");
  let timer: ReturnType<typeof setTimeout> | undefined;
  const email = await Promise.race([
    PostalMime.parse(raw, {
    rfc822Attachments: true,
    attachmentEncoding: "arraybuffer",
    maxNestingDepth: 32,
    maxHeadersSize: 512 * 1024,
    }),
    new Promise<never>((_resolve, reject) => { timer = setTimeout(() => reject(new Error("Email parsing exceeded the 30-second safety limit.")), EMAIL_LIMITS.milliseconds); }),
  ]).finally(() => { if (timer) clearTimeout(timer); });
  const headers = new Map(email.headers.map((header) => [header.key.toLowerCase(), header.value]));
  const body = (email.text?.trim() || htmlToText(email.html || "") || "(No message body supplied)")
    .replace(/\r\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
  if (new TextEncoder().encode(body).byteLength > EMAIL_LIMITS.bodyBytes) throw new Error("Email body exceeds the 16 MB extracted-text safety limit.");
  if (email.attachments.length > EMAIL_LIMITS.attachments) throw new Error("Email contains more than 500 attachments and cannot be processed safely.");
  return {
    headers,
    body,
    attachments: email.attachments.map((attachment) => {
      const raw = attachment.content;
      const content = raw instanceof Uint8Array
        ? raw
        : raw instanceof ArrayBuffer
          ? new Uint8Array(raw)
          : typeof raw === "string"
            ? new TextEncoder().encode(raw)
            : new Uint8Array();
      return {
        name: attachment.filename || "Unnamed attachment",
        mimeType: attachment.mimeType || "application/octet-stream",
        disposition: attachment.disposition || "attachment",
        content,
      };
    }),
  };
}
