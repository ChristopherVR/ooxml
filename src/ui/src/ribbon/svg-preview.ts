/** Escape text for an SVG text node or attribute. */
export function escapeSvgText(text: string): string {
	return text
		.replace(/&/gu, '&amp;')
		.replace(/</gu, '&lt;')
		.replace(/>/gu, '&gt;')
		.replace(/"/gu, '&quot;');
}

/** A colour safe to put in an attribute: hex, rgb()/rgba(), or `none`. */
export function safeColor(color: string | undefined, fallback = 'none'): string {
	if (!color) {
		return fallback;
	}
	const c = color.trim();
	if (/^#[0-9a-f]{3,8}$/iu.test(c) || /^rgba?\([\d\s.,%]+\)$/iu.test(c) || c === 'none') {
		return c;
	}
	if (c === 'transparent') {
		return 'none';
	}
	return fallback;
}

/** The shared shell every tile uses. */
export function svgTile(width: number, height: number, defs: string, body: string): string {
	return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" aria-hidden="true" focusable="false">${defs ? `<defs>${defs}</defs>` : ''}${body}</svg>`;
}
