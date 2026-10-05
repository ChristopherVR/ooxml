import type { CellRange } from '../address.js';
import type {
	CellStyle,
	CellValue,
	Color,
	HorizontalAlignment,
	VerticalAlignment,
} from '../model.js';
import { type StylePatch, patchStyle } from '../styles.js';
import { parseCellInput } from './deps.js';
import type { ClipboardCell, ClipboardCells } from './types.js';

const NAMED_COLORS: Record<string, string> = {
	black: '000000',
	white: 'FFFFFF',
	red: 'FF0000',
	green: '008000',
	lime: '00FF00',
	blue: '0000FF',
	yellow: 'FFFF00',
	orange: 'FFA500',
	purple: '800080',
	gray: '808080',
	grey: '808080',
	silver: 'C0C0C0',
	navy: '000080',
	maroon: '800000',
	teal: '008080',
	olive: '808000',
	aqua: '00FFFF',
	fuchsia: 'FF00FF',
};

/** A CSS colour (`#rgb`, `#rrggbb`, `rgb()`, a basic name) as a model colour. */
export function parseCssColor(text: string): Color | undefined {
	const value = text.trim().toLowerCase();
	const hex = /^#([0-9a-f]{3}|[0-9a-f]{6})$/.exec(value)?.[1];
	if (hex)
		return {
			rgb: (hex.length === 3 ? [...hex].map((ch) => ch + ch).join('') : hex).toUpperCase(),
		};
	const rgb = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/.exec(value);
	if (rgb)
		return {
			rgb: rgb
				.slice(1, 4)
				.map((n) => Math.min(255, Number(n)).toString(16).padStart(2, '0'))
				.join('')
				.toUpperCase(),
		};
	const named = NAMED_COLORS[value];
	return named ? { rgb: named } : undefined;
}

const ENTITIES: Record<string, string> = {
	amp: '&',
	lt: '<',
	gt: '>',
	quot: '"',
	apos: "'",
	nbsp: ' ',
};

export function decodeEntities(text: string): string {
	return text.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (whole, name: string) => {
		if (name.startsWith('#x') || name.startsWith('#X'))
			return String.fromCodePoint(parseInt(name.slice(2), 16));
		if (name.startsWith('#')) return String.fromCodePoint(Number(name.slice(1)));
		return ENTITIES[name.toLowerCase()] ?? whole;
	});
}

function parseAttributes(text: string): Map<string, string> {
	const attrs = new Map<string, string>();
	for (const m of text.matchAll(/([\w:-]+)\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+))/g))
		attrs.set((m[1] ?? '').toLowerCase(), decodeEntities(m[2] ?? m[3] ?? m[4] ?? ''));
	return attrs;
}

function parseCss(text: string): Map<string, string> {
	const out = new Map<string, string>();
	for (const decl of text.split(';')) {
		const colon = decl.indexOf(':');
		if (colon < 0) continue;
		out.set(decl.slice(0, colon).trim().toLowerCase(), decl.slice(colon + 1).trim());
	}
	return out;
}

/** Excel writes `mso-number-format:"0\.00"` with `\` escapes and `\0022` for quotes. */
function decodeNumberFormat(text: string): string {
	return text
		.replace(/^["']|["']$/g, '')
		.replace(/\\([0-9A-Fa-f]{4})/g, (_m, hex: string) => String.fromCharCode(parseInt(hex, 16)))
		.replace(/\\(.)/g, '$1');
}

const H_ALIGN: Record<string, HorizontalAlignment> = {
	left: 'left',
	center: 'center',
	right: 'right',
	justify: 'justify',
};
const V_ALIGN: Record<string, VerticalAlignment> = {
	top: 'top',
	middle: 'center',
	bottom: 'bottom',
};

function stylePatchFromCss(css: Map<string, string>, inner: string): StylePatch {
	const patch: StylePatch = {};
	const font: NonNullable<StylePatch['font']> = {};
	const weight = css.get('font-weight');
	if ((weight && (weight === 'bold' || Number(weight) >= 600)) || /<(b|strong)\b/i.test(inner))
		font.bold = true;
	if (css.get('font-style') === 'italic' || /<(i|em)\b/i.test(inner)) font.italic = true;
	const decoration = css.get('text-decoration') ?? '';
	if (decoration.includes('underline') || /<u\b/i.test(inner)) font.underline = 'single';
	if (decoration.includes('line-through') || /<(s|strike)\b/i.test(inner)) font.strike = true;
	const color = css.get('color');
	const fontColor = color && parseCssColor(color);
	if (fontColor) font.color = fontColor;
	const size = /^([\d.]+)(pt|px)$/.exec(css.get('font-size') ?? '');
	if (size) font.size = Number(size[1]) * (size[2] === 'px' ? 0.75 : 1);
	const family = css
		.get('font-family')
		?.split(',')[0]
		?.trim()
		.replace(/^["']|["']$/g, '');
	if (family) font.name = family;
	if (Object.keys(font).length) patch.font = font;
	const background = css.get('background-color') ?? css.get('background');
	const fill = background && parseCssColor(background.split(/\s+/)[0] ?? '');
	if (fill) patch.fill = { type: 'pattern', pattern: 'solid', fgColor: fill };
	const h = H_ALIGN[css.get('text-align') ?? ''];
	const v = V_ALIGN[css.get('vertical-align') ?? ''];
	if (h || v) patch.alignment = { ...(h ? { horizontal: h } : {}), ...(v ? { vertical: v } : {}) };
	const numFmt = css.get('mso-number-format');
	if (numFmt) patch.numFmt = decodeNumberFormat(numFmt);
	return patch;
}

function cellText(inner: string): string {
	return decodeEntities(
		inner
			.replace(/<!--[\s\S]*?-->/g, '')
			.replace(/<br\s*\/?>/gi, '\u0000')
			.replace(/<[^>]+>/g, '')
			.replace(/\s+/g, ' ')
			.replace(/ ?\u0000 ?/g, '\n'),
	)
		.replace(/ /g, ' ')
		.trim();
}

/**
 * Parses the first `<table>` of clipboard HTML (as Excel, Sheets or a browser write it) into
 * cells. Class rules in `<style>` and inline styles become formats over `base`; `x:num` and
 * `mso-number-format` are honoured; row and column spans become merges.
 */
export function parseHtmlTable(html: string, base: CellStyle): ClipboardCells | undefined {
	const table = /<table\b[^>]*>([\s\S]*?)<\/table>/i.exec(html)?.[1];
	if (table === undefined) return undefined;
	const classes = new Map<string, string>();
	for (const block of html.matchAll(/<style\b[^>]*>([\s\S]*?)<\/style>/gi))
		for (const rule of (block[1] ?? '')
			.replace(/<!--|-->/g, '')
			.matchAll(/\.([\w-]+)\s*\{([^}]*)\}/g))
			classes.set(rule[1] ?? '', rule[2] ?? '');
	const grid: (ClipboardCell | null)[][] = [];
	const taken = new Set<string>();
	const merges: CellRange[] = [];
	let r = 0;
	for (const tr of table.matchAll(/<tr\b[^>]*>([\s\S]*?)(?=<tr\b|<\/tbody|<\/table|$)/gi)) {
		const row = (grid[r] ??= []);
		let c = 0;
		for (const td of (tr[1] ?? '').matchAll(
			/<t([dh])\b([^>]*)>([\s\S]*?)(?:<\/t[dh]>|(?=<t[dh]\b)|$)/gi,
		)) {
			while (taken.has(`${r},${c}`)) c++;
			const attrs = parseAttributes(td[2] ?? '');
			const inner = td[3] ?? '';
			const css = new Map<string, string>();
			for (const name of (attrs.get('class') ?? '').split(/\s+/))
				for (const [k, v] of parseCss(classes.get(name) ?? '')) css.set(k, v);
			for (const [k, v] of parseCss(attrs.get('style') ?? '')) css.set(k, v);
			const patch = stylePatchFromCss(css, inner);
			if (td[1]?.toLowerCase() === 'h') patch.font = { bold: true, ...patch.font };
			const text = cellText(inner);
			row[c] = text === '' && !Object.keys(patch).length ? null : toCell(text, attrs, patch, base);
			const rowSpan = Math.max(1, Number(attrs.get('rowspan') ?? 1) || 1);
			const colSpan = Math.max(1, Number(attrs.get('colspan') ?? 1) || 1);
			if (rowSpan > 1 || colSpan > 1)
				merges.push({
					start: { row: r, col: c },
					end: { row: r + rowSpan - 1, col: c + colSpan - 1 },
				});
			for (let dr = 0; dr < rowSpan; dr++)
				for (let dc = 0; dc < colSpan; dc++) {
					taken.add(`${r + dr},${c + dc}`);
					if (dr || dc) (grid[r + dr] ??= [])[c + dc] = null;
				}
			c += colSpan;
		}
		r++;
	}
	const rows = grid.length;
	const cols = Math.max(0, ...grid.map((g) => g.length));
	const data = Array.from({ length: rows }, (_v, i) =>
		Array.from({ length: cols }, (_w, j) => grid[i]?.[j] ?? null),
	);
	return rows && cols ? { rows, cols, data, merges } : undefined;
}

function toCell(
	text: string,
	attrs: Map<string, string>,
	patch: StylePatch,
	base: CellStyle,
): ClipboardCell {
	let value: CellValue = text;
	const num = attrs.get('x:num');
	if (num !== undefined && num !== '' && Number.isFinite(Number(num))) value = Number(num);
	else if (attrs.has('x:bool')) value = text.toUpperCase() === 'TRUE';
	else if (patch.numFmt !== '@' && text !== '') {
		const parsed = parseCellInput(text);
		if (parsed.formula === undefined) value = parsed.value;
		if (!patch.numFmt && parsed.numFmt && typeof value === 'number') patch.numFmt = parsed.numFmt;
	}
	return { value, text, style: patchStyle(base, patch) };
}
