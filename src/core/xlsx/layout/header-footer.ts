import type { PageSetup } from '../model';

export type HeaderFooterSection = 'left' | 'center' | 'right';
export type HeaderFooterField = 'page' | 'pages' | 'sheet' | 'file' | 'date' | 'time';
export type HeaderFooterToken =
	| { type: 'text'; value: string }
	| { type: 'field'; field: HeaderFooterField }
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
	color?: string;
}
export interface HeaderFooterRun {
	text: string;
	style: HeaderFooterStyle;
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
			if (offset) {
				const raw = value.slice(i, i + 2 + offset.length);
				result.sections[section].push({ type: 'unsupported', raw });
				result.unsupported.push(raw);
			} else result.sections[section].push({ type: 'field', field });
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
		if (token.field === 'page') return String(context.page);
		if (token.field === 'pages') return String(context.pages);
		if (token.field === 'sheet') return context.sheet;
		if (token.field === 'file') return context.file;
		return new Intl.DateTimeFormat(
			context.locale,
			token.field === 'date' ? { dateStyle: 'short' } : { hour: 'numeric', minute: '2-digit' },
		).format(date);
	};
	const render = (section: HeaderFooterSection) => {
		const runs: HeaderFooterRun[] = [];
		let style: HeaderFooterStyle = {};
		for (const token of parsed.sections[section]) {
			if (token.type === 'format') {
				const toggle = /^&[BIUS]$/.test(token.raw);
				style = { ...style, [token.property]: toggle ? !style[token.property] : token.value };
			} else if (token.type !== 'unsupported') {
				const text = token.type === 'text' ? token.value : field(token);
				const previous = runs.at(-1);
				if (text && previous && JSON.stringify(previous.style) === JSON.stringify(style))
					previous.text += text;
				else if (text) runs.push({ text, style: { ...style } });
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
