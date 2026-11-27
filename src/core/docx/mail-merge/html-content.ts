// HTML merge values (`mergeFields`'s `html` names) read into document content, DOM-free. Only an
// allow-list is read: paragraphs and headings, bold, italic, underline, strikethrough, super- and
// subscript, colour, line breaks, links (http, https, mailto and in-document anchors), lists and
// tables. Every other element gives only its text; script and style content, comments and
// pictures are dropped, and nothing is ever fetched.
import { parseCssDeclarations, tokenizeHtml } from '../../text/html-tokens';
import type { Block, Paragraph, Table, TableCell, TextRun } from '../model';
import { DEFAULT_TABLE_BORDERS } from '../table-defaults';
import {
	BLOCKS,
	MAX_HTML_TOKENS,
	tagFormat,
	VOID,
	type HtmlMergeIssue,
	type RunFormat,
} from './html-format';

/** Deeper element nesting is read as text. */
const MAX_DEPTH = 64;

export type { HtmlMergeIssue } from './html-format';

export interface HtmlMergeContext {
	/** A new paragraph or table id. */
	id(): string;
	/** A new list of the given kind; its `numId`. */
	list(kind: 'bullet' | 'decimal'): number;
	/** The paragraph style id of heading `level` (1 to 6), when the document has one. */
	headingStyle(level: number): string | undefined;
}

export interface HtmlMergeContent {
	/** True when the value holds no block element: its runs replace the field in place. */
	inline: boolean;
	/** The runs of an inline value, formatted relative to the field's run. */
	runs: TextRun[];
	/** The blocks of a block value. */
	blocks: Block[];
	issues: HtmlMergeIssue[];
}

interface ListLevel {
	numId: number;
	level: number;
}

interface TableBuilder {
	table: Table;
	row?: TableCell[] | undefined;
	cell?: TableCell | undefined;
}

/**
 * Reads `html` into content. `base` is the field's run: inline runs and the runs of new
 * paragraphs take its formatting where the HTML sets none.
 */
export function readHtmlMergeValue(
	html: string,
	base: TextRun,
	host: Paragraph,
	context: HtmlMergeContext,
): HtmlMergeContent {
	const tokens = tokenizeHtml(html, MAX_HTML_TOKENS);
	const issues = new Set<HtmlMergeIssue>();
	const inline = !tokens.some((token) => token.kind === 'open' && BLOCKS.has(token.name));
	const blocks: Block[] = [];
	const formats: { name: string; format: RunFormat }[] = [];
	const lists: ListLevel[] = [];
	const tables: TableBuilder[] = [];
	let runs: TextRun[] = [];
	let paragraphProps: Partial<Paragraph> = {};
	let pre = 0;
	const {
		field: _field,
		fieldInstanceId: _id,
		fieldChar: _char,
		fieldCode: _code,
		fieldFlags: _flags,
		...baseRun
	} = base;
	const current = (): RunFormat => Object.assign({}, ...formats.map((entry) => entry.format));
	const addText = (text: string) => {
		const value = pre ? text.replace(/\r\n?/g, '\n') : text.replace(/[\t\r\n ]+/g, ' ');
		if (!value) return;
		const last = runs.at(-1);
		const leading = !pre && (!last || last.text.endsWith(' ') || last.text.endsWith('\n'));
		const trimmed = leading && !inline ? value.replace(/^ /, '') : value;
		if (trimmed) runs.push({ ...baseRun, ...current(), text: trimmed });
	};
	const flush = (force = false) => {
		if (inline) return;
		while (runs.at(-1)?.text.endsWith(' ')) {
			const last = runs.at(-1)!;
			let end = last.text.length;
			while (end > 0 && last.text.charCodeAt(end - 1) === 32) end--;
			const text = last.text.slice(0, end);
			if (text) {
				runs[runs.length - 1] = { ...last, text };
				break;
			}
			runs.pop();
		}
		if (!runs.length && !force) return;
		const paragraph: Paragraph = {
			type: 'paragraph',
			id: context.id(),
			runs: runs.length ? runs : [{ ...baseRun, text: '' }],
			...(host.style !== undefined && { style: host.style }),
			...(host.direction !== undefined && { direction: host.direction }),
			...paragraphProps,
		};
		runs = [];
		const builder = tables[0];
		if (!builder) blocks.push(paragraph);
		else {
			// Text inside a table but outside its cells goes into a cell of its own.
			if (!builder.cell) {
				if (!builder.row) builder.table.rows.push((builder.row = []));
				builder.cell = { paragraphs: [] };
				builder.row.push(builder.cell);
			}
			builder.cell.paragraphs.push(paragraph);
		}
	};
	const closeFormat = (name: string) => {
		const index = formats.map((entry) => entry.name).lastIndexOf(name);
		if (index >= 0) formats.splice(index);
	};
	for (const token of tokens) {
		if (token.kind === 'text') {
			addText(token.text);
			continue;
		}
		if (token.kind === 'conditional') continue;
		const { name } = token;
		if (token.kind === 'open') {
			if (name === 'img') {
				issues.add('image-refused');
				continue;
			}
			if (name === 'br') {
				runs.push({ ...baseRun, ...current(), text: '\n' });
				continue;
			}
			if (VOID.has(name) || token.selfClosing) continue;
			if (formats.length >= MAX_DEPTH) continue;
			if (name === 'p' || name === 'div' || name === 'blockquote' || name === 'section') {
				flush();
				paragraphProps = alignment(token.attrs, paragraphProps);
			} else if (/^h[1-6]$/.test(name)) {
				flush();
				const style = context.headingStyle(Number(name[1]));
				paragraphProps = style ? { style } : {};
				if (!style) {
					formats.push({ name, format: { bold: true } });
					continue;
				}
			} else if (name === 'pre') {
				flush();
				pre++;
			} else if (name === 'ul' || name === 'ol') {
				flush();
				const parent = lists.at(-1);
				lists.push(
					parent
						? { numId: parent.numId, level: Math.min(parent.level + 1, 8) }
						: { numId: context.list(name === 'ul' ? 'bullet' : 'decimal'), level: 0 },
				);
			} else if (name === 'li') {
				flush();
				const list = lists.at(-1);
				paragraphProps = list ? { numbering: { numId: list.numId, level: list.level } } : {};
			} else if (name === 'table') {
				flush();
				if (tables.length) {
					issues.add('nested-table-flattened');
					tables.push({ ...tables.at(-1)!, table: tables.at(-1)!.table });
				} else {
					const border = Number(token.attrs.get('border') ?? '0');
					tables.push({
						table: {
							type: 'table',
							id: context.id(),
							structureEditable: true,
							rows: [],
							...(border > 0 && { borders: DEFAULT_TABLE_BORDERS }),
						},
					});
				}
			} else if (name === 'tr') {
				const builder = tables.at(-1);
				if (builder && tables.length === 1) {
					flush();
					builder.row = [];
					builder.table.rows.push(builder.row);
					builder.cell = undefined;
				}
			} else if (name === 'td' || name === 'th') {
				const builder = tables.at(-1);
				if (builder && tables.length === 1) {
					flush();
					if (!builder.row) {
						builder.row = [];
						builder.table.rows.push(builder.row);
					}
					const span = Math.min(Math.max(Number(token.attrs.get('colspan') ?? '1') | 0, 1), 63);
					builder.cell = { paragraphs: [], ...(span > 1 && { gridSpan: span }) };
					builder.row.push(builder.cell);
				} else if (builder) addText(' ');
			}
			formats.push({ name, format: tagFormat(name, token.attrs, issues) });
			continue;
		}
		// A closing tag.
		if (name === 'p' || name === 'div' || name === 'li' || name === 'blockquote') {
			flush();
			paragraphProps = {};
		} else if (/^h[1-6]$/.test(name)) {
			flush();
			paragraphProps = {};
		} else if (name === 'pre') {
			flush();
			pre = Math.max(0, pre - 1);
		} else if (name === 'ul' || name === 'ol') {
			flush();
			lists.pop();
		} else if (name === 'td' || name === 'th') {
			const builder = tables.at(-1);
			if (builder && tables.length === 1 && builder.cell) {
				flush(!builder.cell.paragraphs.length);
				builder.cell = undefined;
			}
		} else if (name === 'table') {
			const builder = tables.pop();
			if (builder && !tables.length) {
				flush();
				const table = finishTable(builder.table, baseRun, context);
				if (table) blocks.push(table);
			}
		}
		closeFormat(name);
	}
	while (tables.length) {
		const builder = tables.pop()!;
		if (!tables.length) {
			flush();
			const table = finishTable(builder.table, baseRun, context);
			if (table) blocks.push(table);
		}
	}
	flush();
	return { inline, runs: inline ? runs : [], blocks, issues: [...issues] };
}

function alignment(attrs: Map<string, string>, props: Partial<Paragraph>): Partial<Paragraph> {
	const css = parseCssDeclarations(attrs.get('style') ?? '');
	const align = (css.get('text-align') ?? attrs.get('align') ?? '').toLowerCase();
	const { align: _align, ...rest } = props;
	return /^(left|center|right|justify)$/.test(align)
		? { ...rest, align: align as Paragraph['align'] & string }
		: rest;
}

/** A table with every row padded to the widest, or undefined when it has no cells. */
function finishTable(table: Table, baseRun: TextRun, context: HtmlMergeContext): Table | undefined {
	const rows = table.rows.filter((row) => row.length);
	if (!rows.length) return undefined;
	const width = (row: TableCell[]) => row.reduce((sum, cell) => sum + (cell.gridSpan ?? 1), 0);
	const columns = Math.max(...rows.map(width));
	for (const row of rows) {
		while (width(row) < columns) row.push({ paragraphs: [] });
		for (const cell of row)
			if (!cell.paragraphs.length)
				cell.paragraphs.push({
					type: 'paragraph',
					id: context.id(),
					runs: [{ ...baseRun, text: '' }],
				});
	}
	return { ...table, rows };
}
