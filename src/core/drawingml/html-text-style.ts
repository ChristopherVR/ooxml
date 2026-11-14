// CSS and presentational HTML attributes of pasted clipboard markup as DrawingML run and
// paragraph properties (`a:rPr`, `a:pPr`). Only values that map onto the model are read; anything
// else (classes, layout, unsupported units) is ignored rather than guessed.
import { parseCssColorHex, parseCssDeclarations } from '../text/html-tokens';
import type { DrawingBullet, DrawingRunProperties } from './text-model';
import type { DrawingColor } from './types';

const srgb = (value: string): DrawingColor => ({ kind: 'srgb', value, transforms: [] });

/** Generic CSS families carry no typeface a slide can name. */
const GENERIC_FAMILIES = new Set([
	'serif',
	'sans-serif',
	'monospace',
	'cursive',
	'fantasy',
	'system-ui',
	'inherit',
	'initial',
	'unset',
]);

/** The first named family of a `font-family` list. */
export function cssFirstFontFamily(value: string): string | undefined {
	for (const part of value.split(',')) {
		const name = part.trim().replace(/^["']|["']$/g, '');
		if (name && name.length <= 64 && !GENERIC_FAMILIES.has(name.toLowerCase())) return name;
	}
	return undefined;
}

/** `font-size` in points: `pt`, `px` (0.75 pt) and the absolute keywords; relative sizes are not resolved. */
export function cssFontSizePt(value: string): number | undefined {
	const keyword: Record<string, number> = {
		'xx-small': 7,
		'x-small': 7.5,
		small: 10,
		medium: 12,
		large: 13.5,
		'x-large': 18,
		'xx-large': 24,
	};
	const v = value.trim().toLowerCase();
	if (keyword[v] !== undefined) return keyword[v];
	const m = /^(\d{1,4}(?:\.\d{1,3})?)(pt|px)$/.exec(v);
	if (!m) return undefined;
	const pt = Number(m[1]) * (m[2] === 'px' ? 0.75 : 1);
	// DrawingML `sz` runs from 1 to 4000 points.
	return pt >= 1 && pt <= 4000 ? Math.round(pt * 100) / 100 : undefined;
}

/** HTML `<font size="1..7">` in points, as browsers render them. */
const FONT_SIZE_ATTR = [7.5, 10, 12, 13.5, 18, 24, 36];

/** Run properties a tag implies before its `style` attribute applies. */
export function tagRunProperties(
	name: string,
	attrs: Map<string, string>,
): DrawingRunProperties | undefined {
	switch (name) {
		case 'b':
		case 'strong':
			return { bold: true };
		case 'i':
		case 'em':
		case 'cite':
		case 'dfn':
			return { italic: true };
		case 'u':
		case 'ins':
			return { underline: 'sng' };
		case 's':
		case 'strike':
		case 'del':
			return { strike: 'sngStrike' };
		case 'sup':
			return { baseline: 0.3 };
		case 'sub':
			return { baseline: -0.25 };
		case 'mark':
			return { highlight: srgb('FFFF00') };
		case 'code':
		case 'kbd':
		case 'samp':
		case 'tt':
		case 'pre':
			return { latin: { typeface: 'Courier New' } };
		case 'h1':
			return { bold: true, sizePt: 24 };
		case 'h2':
			return { bold: true, sizePt: 18 };
		case 'h3':
			return { bold: true, sizePt: 14 };
		case 'h4':
		case 'h5':
		case 'h6':
			return { bold: true };
		case 'font': {
			const props: DrawingRunProperties = {};
			const face = attrs.get('face');
			const family = face ? cssFirstFontFamily(face) : undefined;
			if (family) props.latin = { typeface: family };
			const color = attrs.get('color');
			const hex = color ? parseCssColorHex(color) : undefined;
			if (hex) props.fill = { kind: 'solid', color: srgb(hex) };
			const size = Number(attrs.get('size'));
			const sizePt = Number.isInteger(size) ? FONT_SIZE_ATTR[size - 1] : undefined;
			if (sizePt !== undefined) props.sizePt = sizePt;
			return props;
		}
		default:
			return undefined;
	}
}

/** Run properties of a `style` attribute. `inline` elements also turn a background into a highlight. */
export function cssRunProperties(style: string, inline: boolean): DrawingRunProperties {
	const css = parseCssDeclarations(style);
	const props: DrawingRunProperties = {};
	const weight = css.get('font-weight');
	if (weight) props.bold = weight === 'bold' || weight === 'bolder' || Number(weight) >= 600;
	const fontStyle = css.get('font-style');
	if (fontStyle) props.italic = fontStyle === 'italic' || fontStyle === 'oblique';
	const decoration = css.get('text-decoration-line') ?? css.get('text-decoration');
	if (decoration) {
		if (decoration.includes('underline')) props.underline = 'sng';
		if (decoration.includes('line-through')) props.strike = 'sngStrike';
		if (decoration.includes('none')) {
			props.underline = 'none';
			props.strike = 'noStrike';
		}
	}
	const family = css.get('font-family');
	const typeface = family ? cssFirstFontFamily(family) : undefined;
	if (typeface) props.latin = { typeface };
	const size = css.get('font-size');
	const sizePt = size ? cssFontSizePt(size) : undefined;
	if (sizePt !== undefined) props.sizePt = sizePt;
	const color = css.get('color');
	const hex = color ? parseCssColorHex(color) : undefined;
	if (hex) props.fill = { kind: 'solid', color: srgb(hex) };
	// Word writes `mso-highlight:yellow` beside the background it paints.
	const shorthand = css.get('background');
	const highlight =
		css.get('mso-highlight') ??
		(inline
			? (css.get('background-color') ??
				(shorthand ? /^(rgba?\([^)]*\)|#\w+|[a-z]+)/i.exec(shorthand.trim())?.[1] : undefined))
			: undefined);
	const highlightHex = highlight ? parseCssColorHex(highlight) : undefined;
	if (highlightHex) props.highlight = srgb(highlightHex);
	const vertical = css.get('vertical-align');
	if (vertical === 'super') props.baseline = 0.3;
	else if (vertical === 'sub') props.baseline = -0.25;
	const caps = css.get('font-variant');
	if (caps === 'small-caps') props.caps = 'small';
	if (css.get('text-transform') === 'uppercase') props.caps = 'all';
	return props;
}

const ALIGN: Record<string, string> = {
	left: 'l',
	start: 'l',
	center: 'ctr',
	right: 'r',
	end: 'r',
	justify: 'just',
};

/** Paragraph alignment (`algn`) of a block's `align` attribute or `text-align`. */
export function blockAlignment(name: string, attrs: Map<string, string>): string | undefined {
	if (name === 'center') return 'ctr';
	const css = parseCssDeclarations(attrs.get('style') ?? '');
	return ALIGN[(css.get('text-align') ?? attrs.get('align') ?? '').trim().toLowerCase()];
}

/** Word's `mso-list:l0 level2 lfo1` paragraph style: the list level, zero based. */
export function msoListLevel(attrs: Map<string, string>): number | undefined {
	const css = parseCssDeclarations(attrs.get('style') ?? '');
	const list = css.get('mso-list');
	if (!list || /^\s*ignore\b/i.test(list)) return undefined;
	const level = /level(\d{1,2})/i.exec(list)?.[1];
	return Math.min(8, Math.max(0, (level ? Number(level) : 1) - 1));
}

/** Marker characters kept as typed; anything else (symbol-font code points) becomes a bullet. */
const MARKER_CHARS = new Set(['•', '▪', '●', '○', '■', '–', '-', '*', 'o']);

/** The DrawingML bullet a typed list marker (`1.`, `a)`, `ii.`, a symbol) stands for. */
export function bulletFromMarker(marker: string): DrawingBullet {
	const text = marker.replace(/\s+/g, '');
	const numbered = /^(\(?)(\d{1,4}|[a-z]|[A-Z]|[ivxlc]{1,6}|[IVXLC]{1,6})([.)])$/.exec(text);
	if (numbered) {
		const body = numbered[2] ?? '';
		const suffix = numbered[1] ? 'ParenBoth' : numbered[3] === ')' ? 'ParenR' : 'Period';
		const kind = /^\d+$/.test(body)
			? 'arabic'
			: /^[ivxlc]+$/.test(body) && body !== 'c'
				? 'romanLc'
				: /^[IVXLC]+$/.test(body) && body !== 'C'
					? 'romanUc'
					: /^[a-z]$/.test(body)
						? 'alphaLc'
						: 'alphaUc';
		const startAt = /^\d+$/.test(body) ? Number(body) : undefined;
		return { kind: 'autoNumber', scheme: `${kind}${suffix}`, ...(startAt ? { startAt } : {}) };
	}
	// Word's symbol fonts write private-use or Wingdings code points; PowerPoint's default is a bullet.
	const char = [...text][0] ?? '';
	return { kind: 'char', char: MARKER_CHARS.has(char) ? char : '•' };
}

/** The bullet an `<ol type start>` or `<ul>` item uses. */
export function listBullet(
	ordered: boolean,
	attrs: Map<string, string>,
	level: number,
): DrawingBullet {
	if (!ordered) return { kind: 'char', char: ['•', '–', '•'][level % 3] ?? '•' };
	const type = attrs.get('type') ?? '1';
	const scheme =
		{ a: 'alphaLcPeriod', A: 'alphaUcPeriod', i: 'romanLcPeriod', I: 'romanUcPeriod' }[type] ??
		'arabicPeriod';
	const start = Math.trunc(Number(attrs.get('start')));
	return {
		kind: 'autoNumber',
		scheme,
		...(Number.isFinite(start) && start > 1 && start < 32768 ? { startAt: start } : {}),
	};
}
