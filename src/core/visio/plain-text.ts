/** Native Visio stores one structural paragraph terminator after nonempty shape text. */
export function encodeVisioPlainText(text: string): string {
	return text ? `${text}\n` : '';
}

/** Text without a stored terminator remains valid; never trim caller whitespace. */
export function decodeVisioPlainText(text: string): string {
	return text.endsWith('\n') ? text.slice(0, -1) : text;
}
