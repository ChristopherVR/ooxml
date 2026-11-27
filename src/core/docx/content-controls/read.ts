// Reading the values of a document's tagged content controls back: each control's content as text
// or as HTML (without fonts), the first control of a tag winning. A control around others reads
// through to them (a group's content holds its controls' values), table cells included; a control
// showing its placeholder reads as empty.
import type { ContentControl } from '../content-control-model';
import { mapBlockLists } from '../mail-merge/stories';
import type { Block, DocumentModel, Table, TableCell, TextRun } from '../model';
import type { TableRowProperties } from '../table-model';
import { chainOf, groupAt, runGroups } from './chains';
import { blocksText, cellBlocks } from './controls';
import { blocksHtml, runsHtml } from './html-out';
import type { ReadContentControlsOptions } from './types';

interface Reader {
	model: DocumentModel;
	options: ReadContentControlsOptions;
	values: Map<string, string>;
}

const runsText = (runs: readonly TextRun[]) =>
	runs
		.filter((run) => !run.fieldChar && run.fieldCode === undefined)
		.map((run) => run.text)
		.join('');

/** The value a control of a state kind reads as: what filling it takes. */
function stateValue(control: ContentControl, text: string): string | undefined {
	if (control.kind === 'checkbox') return String(Boolean(control.checked));
	if (control.kind === 'date' && control.fullDate) return control.fullDate.slice(0, 10);
	if (control.kind === 'dropDownList' || control.kind === 'comboBox')
		return control.items?.find((item) => item.display === text)?.value ?? text;
	return undefined;
}

function record(
	reader: Reader,
	control: ContentControl,
	text: () => string,
	html: () => string,
): void {
	const tag = control.tag;
	if (tag === undefined || reader.values.has(tag)) return;
	if (reader.options.prefix !== undefined && !tag.startsWith(reader.options.prefix)) return;
	if (control.showingPlaceholder) {
		reader.values.set(tag, '');
		return;
	}
	const state = stateValue(control, text());
	reader.values.set(tag, state ?? (reader.options.as === 'html' ? html() : text()));
}

function readBlocks(blocks: readonly Block[], depth: number, reader: Reader): void {
	for (const group of groupAt(blocks, depth, chainOf)) {
		if (group.control) {
			const items = group.items;
			record(
				reader,
				group.control,
				() => blocksText(items),
				() => blocksHtml(items, reader.model),
			);
			readBlocks(items, depth + 1, reader);
			continue;
		}
		for (const block of group.items) {
			if (block.type === 'table') readTable(block, reader);
			else
				for (const runs of runGroups(block))
					if (runs.control)
						record(
							reader,
							runs.control,
							() => runsText(runs.runs),
							() => runsHtml(runs.runs),
						);
		}
	}
}

interface Row {
	cells: TableCell[];
	props: TableRowProperties | undefined;
}

function readRows(rows: readonly Row[], depth: number, reader: Reader): void {
	for (const group of groupAt(rows, depth, (row) => chainOf(row.props))) {
		if (group.control) {
			const table: Table = { type: 'table', id: '', rows: group.items.map((row) => row.cells) };
			record(
				reader,
				group.control,
				() => blocksText([table]),
				() => blocksHtml([table], reader.model),
			);
			readRows(group.items, depth + 1, reader);
			continue;
		}
		for (const row of group.items)
			for (const cell of row.cells) {
				const blocks = cellBlocks(cell);
				if (cell.contentControl)
					record(
						reader,
						cell.contentControl,
						() => blocksText(blocks),
						() => blocksHtml(blocks, reader.model),
					);
				readBlocks(blocks, 0, reader);
			}
	}
}

const readTable = (table: Table, reader: Reader) =>
	readRows(
		table.rows.map((cells, index) => ({ cells, props: table.rowProperties?.[index] })),
		0,
		reader,
	);

/**
 * The values of the tagged content controls of `model`, by tag: the body, headers, footers and
 * notes, in reading order. `as: 'html'` reads rich content as HTML without fonts.
 */
export function readContentControls(
	model: DocumentModel,
	options: ReadContentControlsOptions = {},
): Record<string, string> {
	const reader: Reader = { model, options, values: new Map() };
	mapBlockLists(model, (blocks, list) => {
		if (!list.cell && list.story !== 'textBox') readBlocks(blocks, 0, reader);
		return blocks;
	});
	// Built from entries, so that a tag such as `__proto__` is a key like any other.
	return Object.fromEntries(reader.values);
}
