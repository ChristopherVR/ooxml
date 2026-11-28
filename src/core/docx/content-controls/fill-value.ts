// The content a value gives a control: plain text or HTML (through the merge's HTML reader) for a
// rich text or plain text control, a state for a check box, a list or a date. The control's own
// typography wins: new runs take the font and size of its first text run, the HTML adds only
// emphasis, colour, links and structure, and new paragraphs keep the paragraph formatting of the
// control's first paragraph.
import {
	checkboxState,
	selectListItem,
	setControlDate,
	withValue,
	type ControlChange,
} from '../content-control-edit';
import type { ContentControl } from '../content-control-model';
import { parseDay } from '../form-fill-values';
import { readHtmlMergeValue, type HtmlMergeContext } from '../mail-merge/html-content';
import type { HtmlMergeIssue } from '../mail-merge/html-format';
import type { Block, Paragraph, TextRun } from '../model';

const PLACEHOLDER_STYLE = 'PlaceholderText';
/** Run properties that make a run something other than formatted text. */
const STRUCTURE = [
	'fieldChar',
	'fieldCode',
	'field',
	'formField',
	'fieldInstanceId',
	'fieldFlags',
	'image',
	'equation',
	'noteReference',
	'noteMark',
	'noteSeparator',
	'break',
	'symbol',
	'positionalTab',
	'revision',
	'contentControl',
	'link',
	'comments',
] as const;

/** A text run in the formatting of `source`, without what made `source` more than text. */
export function textRun(source: TextRun | undefined, text: string): TextRun {
	const run = { ...source, text } as TextRun & Record<string, unknown>;
	for (const key of STRUCTURE) delete run[key];
	if (run.style === PLACEHOLDER_STYLE) delete run.style;
	return run;
}

/** The run new content is formatted like: the first one with text, else the first plain run. */
export function baseRun(paragraphs: readonly Paragraph[]): TextRun {
	const runs = paragraphs.flatMap((paragraph) => paragraph.runs);
	const plain = (run: TextRun) => !run.fieldChar && run.fieldCode === undefined && !run.image;
	return textRun(runs.find((run) => plain(run) && run.text) ?? runs.find(plain), '');
}

/** Paragraph properties that belong to the paragraph itself, not to its formatting. */
const OWN = [
	'type',
	'id',
	'runs',
	'bookmarks',
	'permissions',
	'markRevision',
	'markInsertion',
	'formatRevision',
	'sourceParagraphPropertiesXml',
	'restoredParagraphPropertiesXml',
	'contentControl',
	'outerContentControls',
] as const;

/** The formatting of `host` that new paragraphs of its control take. */
export function paragraphFormatting(host: Paragraph | undefined): Partial<Paragraph> {
	if (!host) return {};
	const formatting = { ...host } as Partial<Paragraph> & Record<string, unknown>;
	for (const key of OWN) delete formatting[key];
	return formatting;
}

export interface ValueContent {
	/** For an inline control. */
	runs: TextRun[];
	/** For a block or cell control. */
	blocks: Block[];
	/** Shown nothing: only spaces, `&nbsp;` or tags without text. */
	empty: boolean;
}

export interface ValueEnvironment {
	html: HtmlMergeContext;
	issues: Set<HtmlMergeIssue>;
	/** The blocks go in a table cell, which cannot hold a table here. */
	cell: boolean;
}

const blank = (text: string) => !text.replace(/ /g, ' ').trim();
const LINE = /\r\n?|\n/;

/** Tables in a cell become one paragraph per row, cells apart by tabs. */
function flattenTables(blocks: Block[], env: ValueEnvironment): Paragraph[] {
	return blocks.flatMap((block): Paragraph[] => {
		if (block.type === 'paragraph') return [block];
		env.issues.add('nested-table-flattened');
		return block.rows.map((row) => ({
			type: 'paragraph',
			id: env.html.id(),
			runs: row.flatMap((cell, index) => [
				...(index ? [{ text: '\t' }] : []),
				...cell.paragraphs.flatMap((paragraph, at) => [
					...(at ? [{ text: ' ' }] : []),
					...paragraph.runs,
				]),
			]),
		}));
	});
}

/** The content of `value` (HTML when `html`) for a control whose content is `paragraphs`. */
export function valueContent(
	value: string,
	html: boolean,
	paragraphs: readonly Paragraph[],
	env: ValueEnvironment,
): ValueContent {
	const base = baseRun(paragraphs);
	const host = paragraphs[0] ?? { type: 'paragraph', id: env.html.id(), runs: [base] };
	let blocks: Block[];
	if (html) {
		const content = readHtmlMergeValue(value, base, host, env.html);
		for (const issue of content.issues) env.issues.add(issue);
		blocks = content.inline
			? [{ type: 'paragraph', id: env.html.id(), runs: content.runs }]
			: content.blocks;
	} else
		blocks = value
			.split(LINE)
			.map((line) => ({ type: 'paragraph', id: env.html.id(), runs: [textRun(base, line)] }));
	if (env.cell) blocks = flattenTables(blocks, env);
	const text = (block: Block): string =>
		block.type === 'paragraph'
			? block.runs.map((run) => run.text).join('')
			: block.rows
					.flat()
					.flatMap((cell) => cell.paragraphs.map(text))
					.join('');
	const empty = blocks.every((block) => blank(text(block)));
	if (empty || !blocks.length) blocks = [{ type: 'paragraph', id: host.id, runs: [base] }];
	const formatting = paragraphFormatting(host);
	blocks = blocks.map((block, index): Block => {
		if (block.type === 'table') return block;
		const own = Object.fromEntries(
			Object.entries(block).filter(([, item]) => item !== undefined),
		) as unknown as Paragraph;
		// The first paragraph takes the host's place, keeping its properties and identity.
		return index === 0
			? { ...host, ...own, id: host.id }
			: ({ ...formatting, ...own } as Paragraph);
	});
	if (blocks[0]?.type === 'paragraph' && blocks[0].id !== host.id) blocks[0].id = host.id;
	const runs = empty
		? [{ ...base }]
		: blocks.flatMap((block, index) => {
				const lines =
					block.type === 'paragraph'
						? block.runs
						: flattenTables([block], env).flatMap((p) => p.runs);
				return index ? [{ ...base, text: '\n' }, ...lines] : lines;
			});
	return { runs, blocks, empty };
}

const TRUE = /^(true|1|yes|x|checked|on)$/i;
const FALSE = /^(false|0|no|unchecked|off|)$/i;

/** The state and text a check box, list or date control takes for `value`, or why it cannot. */
export function stateChange(control: ContentControl, value: string): ControlChange | string {
	const trimmed = value.trim();
	if (control.kind === 'checkbox') {
		if (TRUE.test(trimmed)) return checkboxState(control, true);
		if (FALSE.test(trimmed)) return checkboxState(control, false);
		return `"${value}" is not a check box state (true or false)`;
	}
	if (control.kind === 'dropDownList' || control.kind === 'comboBox') {
		const items = control.items ?? [];
		let index = items.findIndex((item) => item.value === trimmed);
		if (index < 0) index = items.findIndex((item) => item.display === trimmed);
		const picked = index >= 0 ? selectListItem(control, index) : undefined;
		if (picked) return picked;
		if (control.kind === 'comboBox') return { control: withValue(control), text: value };
		return `"${value}" is none of the drop-down list's entries`;
	}
	if (control.kind === 'date') {
		const date = parseDay(trimmed);
		return date ? setControlDate(control, date) : `"${value}" is not a date (YYYY-MM-DD)`;
	}
	return { control: withValue(control), text: value };
}
