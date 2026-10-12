import type { PageSetup } from '../model';

export type HeaderFooterSection = 'left' | 'center' | 'right';
export type HeaderFooterField = 'page' | 'pages' | 'sheet' | 'file' | 'date' | 'time';
export type HeaderFooterToken =
	| { type: 'text'; value: string }
	| { type: 'field'; field: HeaderFooterField }
	| { type: 'unsupported'; raw: string };
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

export function renderHeaderFooter(
	value: string,
	context: HeaderFooterContext,
): { sections: Record<HeaderFooterSection, string>; unsupported: string[] } {
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
	const render = (section: HeaderFooterSection) =>
		parsed.sections[section]
			.map((token) =>
				token.type === 'text' ? token.value : token.type === 'field' ? field(token) : '',
			)
			.join('');
	return {
		sections: { left: render('left'), center: render('center'), right: render('right') },
		unsupported: parsed.unsupported,
	};
}
