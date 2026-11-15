/** Finding Word's clipboard HTML and the part of it that was copied. */

/**
 * Whether `html` came from Word (or another Office program that writes Word's HTML, such as
 * Outlook): its namespace, its `ProgId`/`Generator` meta, or its `Mso` classes and `mso-` CSS.
 */
export function isWordClipboardHtml(html: string): boolean {
	return (
		/urn:schemas-microsoft-com:office:word/i.test(html) ||
		hasWordMeta(html) ||
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
	const body = bodyContent(html);
	if (body !== undefined) return body;
	// A raw clipboard header (`Version:1.0` and offsets) precedes the markup.
	return /^Version:\d/.test(html) ? html.slice(Math.max(0, html.indexOf('<'))) : html;
}

/** Whether a `<meta>` tag names Word as the producer (`ProgId` or `Generator`), scanned linearly. */
function hasWordMeta(html: string): boolean {
	const lower = html.toLowerCase();
	for (let at = lower.indexOf('<meta'); at >= 0; at = lower.indexOf('<meta', at + 5)) {
		const close = lower.indexOf('>', at);
		if (close < 0) break;
		const tag = lower.slice(at, close);
		const content = tag.indexOf('content=');
		if (content < 0) continue;
		const value = tag.slice(content + 8).replace(/^["']/, '');
		if (value.startsWith('word.document') || value.startsWith('microsoft word')) return true;
	}
	return false;
}

/** What lies inside `<body>` (to `</body>` or the end), or undefined without one. */
function bodyContent(html: string): string | undefined {
	const lower = html.toLowerCase();
	let at = lower.indexOf('<body');
	// Skip tags that only start with "body", such as a custom `<bodyx>`.
	while (at >= 0 && /[a-z0-9-]/.test(lower[at + 5] ?? '')) at = lower.indexOf('<body', at + 5);
	if (at < 0) return undefined;
	const open = lower.indexOf('>', at);
	if (open < 0) return undefined;
	const close = lower.indexOf('</body', open);
	return html.slice(open + 1, close < 0 ? html.length : close);
}
