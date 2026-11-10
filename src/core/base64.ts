const BASE64 = /^[A-Za-z0-9+/]*={0,2}$/;

/**
 * Normalise a Base64 string returned by Outlook. Returns undefined when the
 * value is not well-formed Base64, so a corrupt read is never attached.
 */
export function normalizeBase64(value: string): string | undefined {
  const compact = /\s/.test(value) ? value.replace(/\s+/g, "") : value;
  if (compact.length % 4 !== 0) return undefined;
  if (!BASE64.test(compact)) return undefined;
  return compact;
}

/** Number of bytes a well-formed Base64 string decodes to. */
export function decodedByteLength(base64: string): number {
  if (base64.length === 0) return 0;
  const padding = base64.endsWith("==") ? 2 : base64.endsWith("=") ? 1 : 0;
  return (base64.length / 4) * 3 - padding;
}
