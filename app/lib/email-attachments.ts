import { parseBundleEmail, type ParsedBundleEmail } from "./email.ts";
import type { WorkbookAnalysis } from "./xlsx.ts";

export const EMAIL_CHILD_LIMITS = {
  depth: 3,
  children: 500,
  totalBytes: 128 * 1024 * 1024,
} as const;

export type EmailChildDisposition = "print-with-email" | "add-as-exhibit" | "leave-out";

export const SUPPORTED_EMAIL_CHILD_EXTENSIONS = new Set(["pdf", "docx", "eml", "xlsx", "txt"]);

export type EmailAttachmentChild = {
  identity: string;
  ordinal: string;
  name: string;
  mimeType: string;
  extension: string;
  sha256: string;
  parentSha256: string;
  file: File;
  supported: boolean;
  nested: boolean;
  workbook?: WorkbookAnalysis;
  sheetSelections?: Array<{ name: string; included: boolean; range: string }>;
  pageSizes?: Array<{
    width: number;
    height: number;
    isA4: boolean;
    wouldAddMarginsOnA4: boolean;
    hasAnnotations: boolean;
  }>;
};

async function sha256Hex(bytes: Uint8Array) {
  const digest = await crypto.subtle.digest("SHA-256", bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer);
  return Array.from(new Uint8Array(digest)).map((part) => part.toString(16).padStart(2, "0")).join("");
}

export function emailChildIdentity(parentSha256: string, ordinal: string, childSha256: string) {
  return `${parentSha256}:${ordinal}:${childSha256}`;
}

export function safeAttachmentName(name: string, ordinal: string) {
  const base = name.replace(/\\/g, "/").split("/").pop()?.trim() || "Unnamed attachment";
  const cleaned = base.replace(/[<>:"|?*\u0000-\u001f]/g, "_").replace(/^\.+/g, "_").slice(0, 180).trim();
  return cleaned || `attachment-${ordinal}`;
}

export function extensionFromAttachment(name: string, mimeType: string) {
  const fromName = name.split(".").pop()?.toLowerCase() ?? "";
  if (fromName && /^[a-z0-9]{1,8}$/.test(fromName)) return fromName;
  const mime = mimeType.toLowerCase();
  if (mime.includes("pdf")) return "pdf";
  if (mime.includes("wordprocessingml") || mime === "application/msword") return "docx";
  if (mime.includes("rfc822") || mime.includes("message/rfc822")) return "eml";
  if (mime.includes("spreadsheetml") || mime.includes("excel")) return "xlsx";
  if (mime.startsWith("text/plain")) return "txt";
  return mime.split("/").pop()?.replace(/[^a-z0-9]/g, "") || "bin";
}

export function isSupportedEmailChild(extension: string) {
  return SUPPORTED_EMAIL_CHILD_EXTENSIONS.has(extension);
}

function attachmentBytes(content: Uint8Array | ArrayBuffer | string | undefined) {
  if (!content) return new Uint8Array();
  if (typeof content === "string") return new TextEncoder().encode(content);
  return content instanceof Uint8Array ? content : new Uint8Array(content);
}

async function collectChildren(
  parsed: ParsedBundleEmail,
  parentSha256: string,
  prefix: string,
  depth: number,
  acc: EmailAttachmentChild[],
  totalBytes: { value: number },
): Promise<void> {
  if (depth > EMAIL_CHILD_LIMITS.depth) return;
  for (const [index, attachment] of parsed.attachments.entries()) {
    if (acc.length >= EMAIL_CHILD_LIMITS.children) {
      throw new Error("Email attachments exceed the 500-child safety limit.");
    }
    const ordinal = prefix ? `${prefix}.${index + 1}` : String(index + 1);
    const bytes = attachmentBytes(attachment.content);
    totalBytes.value += bytes.byteLength;
    if (totalBytes.value > EMAIL_CHILD_LIMITS.totalBytes) {
      throw new Error("Email attachments exceed the 128 MB extracted-child safety limit.");
    }
    const name = safeAttachmentName(attachment.name, ordinal);
    const extension = extensionFromAttachment(name, attachment.mimeType);
    const childSha256 = await sha256Hex(bytes);
    const file = new File([bytes], name, { type: attachment.mimeType || "application/octet-stream" });
    acc.push({
      identity: emailChildIdentity(parentSha256, ordinal, childSha256),
      ordinal,
      name,
      mimeType: attachment.mimeType || "application/octet-stream",
      extension,
      sha256: childSha256,
      parentSha256,
      file,
      supported: isSupportedEmailChild(extension),
      nested: depth > 1,
    });
    if (extension === "eml" && depth < EMAIL_CHILD_LIMITS.depth && bytes.byteLength) {
      try {
        const nested = await parseBundleEmail(new TextDecoder("utf-8", { fatal: false }).decode(bytes));
        await collectChildren(nested, parentSha256, ordinal, depth + 1, acc, totalBytes);
      } catch {
        // Nested EML that cannot be parsed remains listed; its children stay unresolved rather than silently included.
      }
    }
  }
}

export async function extractEmailChildren(parsed: ParsedBundleEmail, parentSha256: string) {
  const children: EmailAttachmentChild[] = [];
  await collectChildren(parsed, parentSha256, "", 1, children, { value: 0 });
  return children;
}

export async function rederiveEmailChildren(parentFile: File, parentSha256: string) {
  const parsed = await parseBundleEmail(await parentFile.text());
  const children = await extractEmailChildren(parsed, parentSha256);
  for (const child of children) {
    const expected = emailChildIdentity(parentSha256, child.ordinal, child.sha256);
    if (child.identity !== expected) throw new Error(`Attachment ${child.name} failed hash-bound identity verification.`);
  }
  return children;
}

export function unresolvedEmailAttachments(
  children: EmailAttachmentChild[] | undefined,
  dispositions: Record<string, EmailChildDisposition> | undefined,
) {
  return (children ?? []).filter((child) => {
    const disposition = dispositions?.[child.identity];
    if (disposition === "leave-out") return false;
    if (!child.supported) return disposition !== "leave-out";
    return !disposition;
  });
}

export function emailChildrenForDisposition(
  children: EmailAttachmentChild[] | undefined,
  dispositions: Record<string, EmailChildDisposition> | undefined,
  disposition: EmailChildDisposition,
) {
  return (children ?? []).filter((child) => dispositions?.[child.identity] === disposition);
}
