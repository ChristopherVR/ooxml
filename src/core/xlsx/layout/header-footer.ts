import type { PageSetup } from '../model';
import { applyTint } from './tint';

export type HeaderFooterSection = 'left' | 'center' | 'right';
export type HeaderFooterField = 'page' | 'pages' | 'sheet' | 'file' | 'date' | 'time';
export type HeaderFooterToken =
	| { type: 'text'; value: string }
	| { type: 'field'; field: HeaderFooterField; offset?: number }
	| { type: 'picture' }
	| {
			type: 'format';
			property: keyof HeaderFooterStyle;
			value: string | number | boolean;
			raw: string;
	  }
	| { type: 'unsupported'; raw: string };
export interface HeaderFooterStyle {
	font?: string;
	size?: number;
	bold?: boolean;
	italic?: boolean;
	underline?: boolean;
	strike?: boolean;
	/** `&E`. */
	doubleUnderline?: boolean;
	/** `&X` and `&Y`. */
	superscript?: boolean;
	subscript?: boolean;
	color?: string;
	/** `&KTT+SSS`: theme slot (as `Color.theme`) and tint in percent, e.g. `05-025`. */
	themeColor?: string;
}
export interface HeaderFooterRun {
	text: string;
	style: HeaderFooterStyle;
	/** The run is the section's picture (`&G`); `text` is empty. */
	picture?: true;
}
export interface ParsedHeaderFooter {
	sections: Record<HeaderFooterSection, HeaderFooterToken[]>;
	unsupported: string[];
}
export interface HeaderFooterContext {
	page: number;
	pages: number;
	sheet: string;
	file: string;
	date?: Date;
	locale?: string;
	/** Workbook theme colours in `Color.theme` order (`RRGGBB`), to resolve `&KTT+SSS` codes. */
	themeColors?: readonly string[];
}
const FIELDS: Record<string, HeaderFooterField> = {
	P: 'page',
	N: 'pages',
	A: 'sheet',
	F: 'file',
	D: 'date',
	T: 'time',
};

export function pageHeaderFooter(
	setup: PageSetup,
	physicalPage: number,
	pageNumber: number,
): { header: string; footer: string } {
	if (physicalPage === 1 && setup.differentFirst)
		return { header: setup.firstHeader ?? '', footer: setup.firstFooter ?? '' };
	if (pageNumber % 2 === 0 && setup.differentOddEven)
		return { header: setup.evenHeader ?? '', footer: setup.evenFooter ?? '' };
	return { header: setup.header ?? '', footer: setup.footer ?? '' };
}

/** SpreadsheetML section/field codes, escaped ampersands and preserved unsupported control codes. */
export function parseHeaderFooter(value: string): ParsedHeaderFooter {
	const result: ParsedHeaderFooter = {
		sections: { left: [], center: [], right: [] },
		unsupported: [],
	};
	let section: HeaderFooterSection = 'center';
	const literal = (text: string) => {
		const tokens = result.sections[section];
		const previous = tokens.at(-1);
		if (previous?.type === 'text') previous.value += text;
		else tokens.push({ type: 'text', value: text });
	};
	for (let i = 0; i < value.length;) {
		if (value[i] !== '&') {
			literal(value[i++]!);
			continue;
		}
		const code = value[i + 1];
		if (!code) {
			literal('&');
			i++;
			continue;
		}
		if (code === '&') {
			literal('&');
			i += 2;
			continue;
		}
		if (['L', 'C', 'R'].includes(code)) {
			section = code === 'L' ? 'left' : code === 'R' ? 'right' : 'center';
			i += 2;
			continue;
		}
		const format = (
			property: keyof HeaderFooterStyle,
			value: string | number | boolean,
			raw: string,
		) => {
			result.sections[section].push({ type: 'format', property, value, raw });
		};
		const toggles: Record<string, keyof HeaderFooterStyle> = {
			B: 'bold',
			I: 'italic',
			U: 'underline',
			S: 'strike',
			E: 'doubleUnderline',
			X: 'superscript',
			Y: 'subscript',
		};
		if (toggles[code]) {
			format(toggles[code]!, true, value.slice(i, i + 2));
			i += 2;
			continue;
		}
		const size = /^&(\d+)(?![\d.])/.exec(value.slice(i));
		if (size && Number(size[1]) >= 1 && Number(size[1]) <= 409) {
			format('size', Number(size[1]), size[0]);
			i += size[0].length;
			continue;
		}
		const color = /^&K([0-9A-Fa-f]{6})/.exec(value.slice(i));
		if (color) {
			format('color', `#${color[1]}`, color[0]);
			i += color[0].length;
			continue;
		}
		const theme = /^&K(\d{2})([+-])(\d{3})/.exec(value.slice(i));
		if (theme) {
			format('themeColor', `${theme[1]}${theme[2]}${theme[3]}`, theme[0]);
			i += theme[0].length;
			continue;
		}
		if (code === 'G') {
			result.sections[section].push({ type: 'picture' });
			i += 2;
			continue;
		}
		const font = /^&"([^",]+),(Regular|Bold|Italic|Bold Italic)"/.exec(value.slice(i));
		if (font) {
			format('font', font[1]!, font[0]);
			format('bold', font[2]!.includes('Bold'), font[0]);
			format('italic', font[2]!.includes('Italic'), font[0]);
			i += font[0].length;
			continue;
		}
		const field = FIELDS[code];
		if (field) {
			const offset = field === 'page' ? /^[+-]\d+/.exec(value.slice(i + 2))?.[0] : undefined;
			result.sections[section].push(
				offset ? { type: 'field', field, offset: Number(offset) } : { type: 'field', field },
			);
			i += 2 + (offset?.length ?? 0);
			continue;
		}
		const raw =
			/^&\[[^\]]*\]|^&"[^"]*"|^&K(?:[0-9A-Fa-f]{6}|\d{2}[+-]\d{3})|^&\d+/.exec(
				value.slice(i),
			)?.[0] ?? value.slice(i, i + 2);
		result.sections[section].push({ type: 'unsupported', raw });
		result.unsupported.push(raw);
		i += raw.length;
	}
	return result;
}

export function renderHeaderFooterRuns(
	value: string,
	context: HeaderFooterContext,
): { sections: Record<HeaderFooterSection, HeaderFooterRun[]>; unsupported: string[] } {
	const parsed = parseHeaderFooter(value);
	const date = context.date ?? new Date();
	const field = (token: Extract<HeaderFooterToken, { type: 'field' }>) => {
		if (token.field === 'page') {
			// Excel adds the offset; a negative result prints the page number followed by the
			// offset digits (measured: page 1 with `&P-7` prints "17").
			const shifted = context.page + (token.offset ?? 0);
			return shifted >= 0 ? String(shifted) : `${context.page}${Math.abs(token.offset ?? 0)}`;
		}
		if (token.field === 'pages') return String(context.pages);
		if (token.field === 'sheet') return context.sheet;
		if (token.field === 'file') return context.file;
		return new Intl.DateTimeFormat(
			context.locale,
			token.field === 'date' ? { dateStyle: 'short' } : { hour: 'numeric', minute: '2-digit' },
		).format(date);
	};
	const resolved = (style: HeaderFooterStyle): HeaderFooterStyle => {
		const { themeColor, ...rest } = style;
		const match = themeColor ? /^(\d{2})([+-])(\d{3})$/.exec(themeColor) : null;
		const base = match ? context.themeColors?.[Number(match[1])] : undefined;
		if (!match || !base) return { ...rest };
		const tint = (match[2] === '-' ? -1 : 1) * (Number(match[3]) / 100);
		return { ...rest, color: `#${applyTint(base.replace(/^#/, ''), tint).replace(/^#/, '')}` };
	};
	const render = (section: HeaderFooterSection) => {
		const runs: HeaderFooterRun[] = [];
		let style: HeaderFooterStyle = {};
		for (const token of parsed.sections[section]) {
			if (token.type === 'format') {
				const toggle = /^&[BIUSEXY]$/.test(token.raw);
				style = { ...style, [token.property]: toggle ? !style[token.property] : token.value };
				// A later RGB colour replaces a theme colour and the other way round.
				if (token.property === 'color') delete style.themeColor;
				if (token.property === 'themeColor') delete style.color;
			} else if (token.type === 'picture') {
				runs.push({ text: '', style: resolved(style), picture: true });
			} else if (token.type !== 'unsupported') {
				const text = token.type === 'text' ? token.value : field(token);
				const previous = runs.at(-1);
				const now = resolved(style);
				if (
					text &&
					previous &&
					!previous.picture &&
					JSON.stringify(previous.style) === JSON.stringify(now)
				)
					previous.text += text;
				else if (text) runs.push({ text, style: now });
			}
		}
		return runs;
	};
	return {
		sections: { left: render('left'), center: render('center'), right: render('right') },
		unsupported: parsed.unsupported,
	};
}

export function renderHeaderFooter(
	value: string,
	context: HeaderFooterContext,
): { sections: Record<HeaderFooterSection, string>; unsupported: string[] } {
	const result = renderHeaderFooterRuns(value, context);
	const text = (section: HeaderFooterSection) =>
		result.sections[section].map((run) => run.text).join('');
	return {
		sections: { left: text('left'), center: text('center'), right: text('right') },
		unsupported: result.unsupported,
	};
}

/** The id of a header or footer picture in `legacyDrawingHF`: `LH`, `CF`, `RHFIRST`, `LFEVEN`... */
export function headerFooterImageKey(
	kind: 'header' | 'footer',
	section: HeaderFooterSection,
	variant: 'default' | 'first' | 'even' = 'default',
): string {
	const suffix = variant === 'first' ? 'FIRST' : variant === 'even' ? 'EVEN' : '';
	return `${section === 'left' ? 'L' : section === 'center' ? 'C' : 'R'}${kind === 'header' ? 'H' : 'F'}${suffix}`;
}
