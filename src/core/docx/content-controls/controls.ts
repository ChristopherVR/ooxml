// Every content control a list of blocks holds, at every level: the block controls of paragraphs
// and tables (and the ones around them), inline controls on runs, row and cell controls, and the
// controls of tables nested in cells. Mapped (a copy's new ids and tags, locks) or read.
import type { ContentControl } from '../content-control-model';
import type { Block, Paragraph, Table, TableCell, TextRun } from '../model';
import { chainOf, withChain } from './chains';

export type ControlMapper = (control: ContentControl) => ContentControl;

function mapChain<T extends Parameters<typeof chainOf>[0] & object>(
	holder: T,
	fn: ControlMapper,
): T {
	const chain = chainOf(holder);
	return chain.length ? withChain(holder, chain.map(fn)) : holder;
}

const mapRuns = (runs: TextRun[], fn: ControlMapper): TextRun[] =>
	runs.map((run) =>
		run.contentControl ? { ...run, contentControl: fn(run.contentControl) } : run,
	);

export const mapParagraphControls = (paragraph: Paragraph, fn: ControlMapper): Paragraph =>
	mapChain({ ...paragraph, runs: mapRuns(paragraph.runs, fn) }, fn);

function mapCell(cell: TableCell, fn: ControlMapper): TableCell {
	const next: TableCell = {
		...cell,
		paragraphs: cell.paragraphs.map((paragraph) => mapParagraphControls(paragraph, fn)),
		...(cell.contentControl && { contentControl: fn(cell.contentControl) }),
	};
	if (cell.sourceBlocks) {
		const byId = new Map(next.paragraphs.map((paragraph) => [paragraph.id, paragraph]));
		next.sourceBlocks = cell.sourceBlocks.map((block) =>
			block.type === 'paragraph' ? (byId.get(block.id) ?? block) : mapTableControls(block, fn),
		);
	}
	return next;
}

export function mapTableControls(table: Table, fn: ControlMapper): Table {
	return mapChain(
		{
			...table,
			rows: table.rows.map((row) => row.map((cell) => mapCell(cell, fn))),
			...(table.rowProperties && {
				rowProperties: table.rowProperties.map((properties) => mapChain(properties, fn)),
			}),
		},
		fn,
	);
}

/** `blocks` with `fn` applied to every control they hold. */
export const mapControls = (blocks: readonly Block[], fn: ControlMapper): Block[] =>
	blocks.map((block) =>
		block.type === 'paragraph' ? mapParagraphControls(block, fn) : mapTableControls(block, fn),
	);

/** Calls `visit` with every control `blocks` hold (a control spanning blocks, once per block). */
export function forEachControl(
	blocks: readonly Block[],
	visit: (control: ContentControl) => void,
): void {
	mapControls(blocks, (control) => {
		visit(control);
		return control;
	});
}

/** Whether a control shows a value: not its placeholder, and some text (or a checked box). */
export function controlShowsValue(control: ContentControl, text: string): boolean {
	if (control.showingPlaceholder) return false;
	if (control.kind === 'checkbox') return Boolean(control.checked);
	return text.replace(/ /g, ' ').trim() !== '';
}

const runText = (runs: readonly TextRun[]) =>
	runs.map((run) => (run.image || run.equation ? '￼' : run.text)).join('');

/** The text of blocks: paragraphs on lines, cells apart by tabs. */
export function blocksText(blocks: readonly Block[]): string {
	return blocks
		.map((block) =>
			block.type === 'paragraph'
				? runText(block.runs)
				: block.rows
						.map((row) => row.map((cell) => blocksText(cellBlocks(cell))).join('\t'))
						.join('\n'),
		)
		.join('\n');
}

/** A cell's blocks in their order, nested tables included. */
export const cellBlocks = (cell: TableCell): Block[] => cell.sourceBlocks ?? cell.paragraphs;

/**
 * Whether any control in `blocks` shows a value. `outer` is the control holding them, which is
 * not counted (its own value is the caller's).
 */
export function anyControlShowsValue(blocks: readonly Block[], outer?: string): boolean {
	let shows = false;
	const visitParagraph = (paragraph: Paragraph) => {
		const byId = new Map<string, { control: ContentControl; text: string }>();
		for (const run of paragraph.runs) {
			const control = run.contentControl;
			if (!control) continue;
			const entry = byId.get(control.id) ?? { control, text: '' };
			entry.text += run.text;
			byId.set(control.id, entry);
		}
		for (const { control, text } of byId.values()) shows ||= controlShowsValue(control, text);
	};
	const visit = (items: readonly Block[]) => {
		const groups = new Map<string, { control: ContentControl; blocks: Block[] }>();
		for (const block of items) {
			for (const control of chainOf(block)) {
				if (control.id === outer) continue;
				const entry = groups.get(control.id) ?? { control, blocks: [] };
				entry.blocks.push(block);
				groups.set(control.id, entry);
			}
			if (block.type === 'paragraph') visitParagraph(block);
			else
				for (const [index, row] of block.rows.entries()) {
					const properties = block.rowProperties?.[index];
					for (const control of chainOf(properties))
						shows ||= controlShowsValue(
							control,
							row.map((cell) => blocksText(cellBlocks(cell))).join(''),
						);
					for (const cell of row) {
						if (cell.contentControl)
							shows ||= controlShowsValue(cell.contentControl, blocksText(cellBlocks(cell)));
						visit(cellBlocks(cell));
					}
				}
		}
		for (const { control, blocks } of groups.values())
			shows ||= controlShowsValue(control, blocksText(blocks));
	};
	visit(blocks);
	return shows;
}
