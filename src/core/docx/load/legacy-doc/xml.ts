// Small string helpers the legacy .doc converter uses to write WordprocessingML parts.

// Characters XML 1.0 cannot carry (control characters other than tab, line feed and return).
// eslint-disable-next-line no-control-regex
const INVALID_XML = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g;

export function esc(text: string): string {
	return text
		.replace(INVALID_XML, '')
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;');
}

export type AttrValue = string | number | boolean | undefined;

/** ` w:name="value"` pairs for the defined entries (`true` writes nothing, `false` is skipped). */
export function attrs(values: Record<string, AttrValue>): string {
	let out = '';
	for (const [name, value] of Object.entries(values)) {
		if (value === undefined || value === false) continue;
		out += ` ${name}="${esc(String(value))}"`;
	}
	return out;
}

/** A self-closing or wrapping element. */
export function el(name: string, values: Record<string, AttrValue> = {}, inner = ''): string {
	return inner ? `<${name}${attrs(values)}>${inner}</${name}>` : `<${name}${attrs(values)}/>`;
}

export const hex6 = (value: number): string => value.toString(16).padStart(6, '0').toUpperCase();

export const NS = {
	w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
	r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
	wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
	a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
	pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
	wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
	mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
} as const;

export const ROOT_NAMESPACES = Object.entries(NS)
	.map(([prefix, uri]) => `xmlns:${prefix}="${uri}"`)
	.join(' ');

export const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

/** Twips per EMU (1 twip = 635 EMU). */
export const EMU_PER_TWIP = 635;
