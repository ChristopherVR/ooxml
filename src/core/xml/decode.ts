/**
 * The text of an XML part's bytes. ECMA-376 Part 2 lets a part be UTF-8 or UTF-16 (either byte
 * order); the byte order mark, or the first character `<` when there is none, tells them apart
 * (XML 1.0, appendix F.1).
 */
export function xmlEncodingOf(bytes: Uint8Array): 'utf-8' | 'utf-16le' | 'utf-16be' {
	const [first, second] = bytes;
	if ((first === 0xff && second === 0xfe) || (first === 0x3c && second === 0)) return 'utf-16le';
	if ((first === 0xfe && second === 0xff) || (first === 0 && second === 0x3c)) return 'utf-16be';
	return 'utf-8';
}

/**
 * {@link xmlEncodingOf} applied: the decoded text, without its byte order mark. Malformed bytes
 * throw unless `fatal` is false, when they read as U+FFFD (as a lenient loader wants).
 */
export function decodeXmlBytes(
	bytes: Uint8Array,
	{ fatal = true }: { fatal?: boolean } = {},
): string {
	return new TextDecoder(xmlEncodingOf(bytes), { fatal }).decode(bytes);
}
