/** Finding Word's clipboard HTML and the part of it that was copied. */

/**
 * Whether `html` came from Word (or another Office program that writes Word's HTML, such as
 * Outlook): its namespace, its `ProgId`/`Generator` meta, or its `Mso` classes and `mso-` CSS.
 */
export function isWordClipboardHtml(html: string): boolean {
	return (
		/urn:schemas-microsoft-com:office:word/i.test(html) ||
		/<meta[^>]+content=["']?(?:Word\.Document|Microsoft Word)/i.test(html) ||
		(/class=["']?Mso/i.test(html) && /mso-/i.test(html))
	);
}

/**
 * The copied fragment: what lies between `<!--StartFragment-->` and `<!--EndFragment-->` (the
 * HTML clipboard format's markers, which browsers keep), else the body, else everything. The
 * byte offsets in a raw `Version:` header are not trusted: they no longer match once the text
 * has been decoded or its line endings changed.
 */
export function clipboardFragment(html: string): string {
	const start = html.search(/<!--\s*StartFragment\s*-->/i);
	const end = html.search(/<!--\s*EndFragment\s*-->/i);
	if (start >= 0 && end > start) return html.slice(html.indexOf('-->', start) + 3, end);
	const body = /<body[^>]*>([\s\S]*?)(?:<\/body>|$)/i.exec(html);
	if (body) return body[1]!;
	// A raw clipboard header (`Version:1.0` and offsets) precedes the markup.
	return /^Version:\d/.test(html) ? html.slice(Math.max(0, html.indexOf('<'))) : html;
}
