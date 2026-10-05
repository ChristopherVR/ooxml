import { FormulaError, type StructuredRef, type StructuredSpecial } from './ast.js';

const SPECIALS: Record<string, StructuredSpecial> = {
	'#all': '#All',
	'#data': '#Data',
	'#headers': '#Headers',
	'#totals': '#Totals',
	'#this row': '#This Row',
};

/** Removes the `'` escapes of a structured-reference column name. */
const unescape = (text: string): string => text.replace(/'(.)/g, '$1');

/** Splits `[a],[b]:[c]` at top-level commas. */
function splitItems(text: string): string[] {
	const items: string[] = [];
	let depth = 0;
	let current = '';
	for (let i = 0; i < text.length; i++) {
		const ch = text[i] ?? '';
		if (ch === "'") {
			current += ch + (text[i + 1] ?? '');
			i++;
			continue;
		}
		if (ch === '[') depth++;
		if (ch === ']') depth--;
		if (ch === ',' && depth === 0) {
			items.push(current.trim());
			current = '';
			continue;
		}
		current += ch;
	}
	if (current.trim()) items.push(current.trim());
	return items;
}

const strip = (item: string): string =>
	item.startsWith('[') && item.endsWith(']') ? item.slice(1, -1) : item;

/** Parses `Table1[...]` (or a bare `[...]` inside a table) into its parts. */
export function parseStructured(text: string, tableName: string): StructuredRef {
	const open = text.indexOf('[');
	const inner = text.slice(open + 1, -1);
	const out: StructuredRef = { specials: [] };
	if (tableName) out.table = tableName;
	const addColumns = (spec: string): void => {
		const range = /^\[((?:[^\]']|'.)*)\]\s*:\s*\[((?:[^\]']|'.)*)\]$/.exec(spec);
		if (range) {
			out.column = unescape(range[1] ?? '');
			out.column2 = unescape(range[2] ?? '');
			return;
		}
		out.column = unescape(strip(spec));
	};
	const handle = (raw: string): void => {
		let item = raw.trim();
		if (item === '') return;
		if (item.startsWith('@')) {
			out.specials.push('#This Row');
			item = item.slice(1).trim();
			if (item === '') return;
		}
		const special = SPECIALS[strip(item).toLowerCase()];
		if (special) {
			out.specials.push(special);
			return;
		}
		if (out.column !== undefined) throw new FormulaError(`Invalid structured reference ${text}`);
		addColumns(item);
	};
	if (!inner.includes('[')) {
		handle(inner);
		return out;
	}
	for (const item of splitItems(inner)) {
		if (/^\[.*\]\s*:\s*\[.*\]$/.test(item) && !SPECIALS[strip(item).toLowerCase()]) {
			addColumns(item);
			continue;
		}
		handle(item);
	}
	return out;
}
