/** Escape text for an HTML body. */
export const escapeHtml = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

/** Plain text as HTML that reads the same: escaped, with its line breaks kept. */
export const textToHtml = (s: string) => escapeHtml(s).replace(/\r?\n/g, "<br>");
