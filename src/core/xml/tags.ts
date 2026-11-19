/**
 * A forward scan over the tags of XML text, for the few readers that need source order the parsed
 * tree drops (or that run before any parse). Every character is read once, so the scan is linear
 * whatever the input: a regular expression such as `<x\b[^>]*>([\s\S]*?)</x>` instead rescans to
 * the end of the text from each opening tag that is never closed, a self-closing one included.
 *
 * Comments, CDATA sections, processing instructions and declarations are skipped; a `>` inside a
 * quoted attribute value does not end a tag. Text that stops inside markup ends the scan.
 *
 * @module xml/tags
 */

/** One start, end or empty-element tag. */
export interface XmlTag {
	/** The qualified name as written (`a:custGeom`). */
	name: string;
	/** The name without its prefix (`custGeom`). */
	local: string;
	/** The raw text between the name and the tag's end (`/>` or `>`), attributes and spaces. */
	attributes: string;
	/** An end tag (`</a:path>`). */
	closing: boolean;
	/** An empty-element tag (`<a:close/>`). */
	selfClosing: boolean;
}

const isNameChar = (code: number): boolean =>
	(code >= 97 && code <= 122) ||
	(code >= 65 && code <= 90) ||
	(code >= 48 && code <= 58) ||
	code === 95 ||
	code === 45 ||
	code === 46 ||
	code > 127;

/** Index just past the end of the markup that starts with `open` at `from`, or -1. */
function skip(xml: string, from: number, open: string, close: string): number {
	const end = xml.indexOf(close, from + open.length);
	return end < 0 ? -1 : end + close.length;
}

/** The tags of `xml` in document order. */
export function* xmlTags(xml: string): Generator<XmlTag> {
	let index = xml.indexOf('<');
	while (index >= 0) {
		let after: number;
		if (xml.startsWith('<!--', index)) after = skip(xml, index, '<!--', '-->');
		else if (xml.startsWith('<![CDATA[', index)) after = skip(xml, index, '<![CDATA[', ']]>');
		else if (xml.startsWith('<?', index)) after = skip(xml, index, '<?', '?>');
		else if (xml.startsWith('<!', index)) after = skip(xml, index, '<!', '>');
		else {
			const closing = xml.charCodeAt(index + 1) === 47;
			const nameStart = index + (closing ? 2 : 1);
			let nameEnd = nameStart;
			while (nameEnd < xml.length && isNameChar(xml.charCodeAt(nameEnd))) nameEnd++;
			if (nameEnd === nameStart) {
				index = xml.indexOf('<', index + 1);
				continue;
			}
			let cursor = nameEnd;
			let quote = 0;
			for (; cursor < xml.length; cursor++) {
				const code = xml.charCodeAt(cursor);
				if (quote) {
					if (code === quote) quote = 0;
				} else if (code === 34 || code === 39) quote = code;
				else if (code === 62) break;
			}
			if (cursor >= xml.length) return;
			const selfClosing = !closing && xml.charCodeAt(cursor - 1) === 47;
			const name = xml.slice(nameStart, nameEnd);
			yield {
				name,
				local: name.slice(name.indexOf(':') + 1),
				attributes: xml.slice(nameEnd, selfClosing ? cursor - 1 : cursor),
				closing,
				selfClosing,
			};
			after = cursor + 1;
		}
		if (after < 0) return;
		index = xml.indexOf('<', after);
	}
}
