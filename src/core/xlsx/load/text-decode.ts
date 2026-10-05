// Text decoding for CSV and other plain-text spreadsheet input. Excel's "CSV" (not "CSV UTF-8")
// is written in the system ANSI code page, which on Western Windows is Windows-1252, without a
// byte-order mark; decoding such a file as UTF-8 turns every accented letter into U+FFFD.

/**
 * Decodes spreadsheet text bytes: a UTF-16 LE/BE or UTF-8 byte-order mark wins; otherwise the
 * bytes are decoded as UTF-8 when they are valid UTF-8 and as Windows-1252 when they are not.
 * The byte-order mark is not part of the result.
 */
export function decodeSpreadsheetText(bytes: Uint8Array): string {
	if (bytes[0] === 0xff && bytes[1] === 0xfe)
		return new TextDecoder('utf-16le').decode(bytes.subarray(2));
	if (bytes[0] === 0xfe && bytes[1] === 0xff)
		return new TextDecoder('utf-16be').decode(bytes.subarray(2));
	if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf)
		return new TextDecoder('utf-8').decode(bytes.subarray(3));
	try {
		return new TextDecoder('utf-8', { fatal: true }).decode(bytes);
	} catch {
		return new TextDecoder('windows-1252').decode(bytes);
	}
}
