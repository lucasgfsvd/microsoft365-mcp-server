/**
 * Strip what costs context but shows nothing. Marketing mail pads its preview
 * line with zero-width joiners, combining grapheme joiners and soft hyphens —
 * 41% of one real newsletter's text body — and runs of spaces around them.
 */
export function tidyText(text: string): string {
  return text
    .replace(/[\u200B\u200C\u200D\u2060\uFEFF\u034F\u00AD\u180E]/g, "")
    .replace(/[ \t\u00A0]+/g, " ")
    .replace(/ *\n */g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}
