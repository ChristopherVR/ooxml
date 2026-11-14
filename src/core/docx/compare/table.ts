import type { Block, Paragraph, Table, TableCell } from '../model';
import type { TableRowProperties } from '../table-model';
import type { CompareContext } from './context';
import { diffKeys } from './diff';
import { compareStory, wholeStory } from './story';
import { rowKey } from './tokens';

type Row = TableCell[];

/** Below this word overlap two rows of a changed stretch are a deletion and an insertion. */
const ROW_SIMILARITY = 0.25;

function rowWords(row: Row): string[] {
	return row
		.flatMap((cell) => cell.paragraphs)
		.flatMap((paragraph) => paragraph.runs.map((run) => run.text).join(''))
		.join(' ')
		.split(/\s+/u)
		.filter(Boolean);
}

/** Dice coefficient of the two rows' words. */
function similarity(a: Row, b: Row): number {
	const left = rowWords(a);
	const right = rowWords(b);
	if (!left.length && !right.length) return 1;
	const counts = new Map<string, number>();
	for (const word of left) counts.set(word, (counts.get(word) ?? 0) + 1);
	let shared = 0;
	for (const word of right) {
		const count = counts.get(word) ?? 0;
		if (count > 0) {
			shared++;
			counts.set(word, count - 1);
		}
	}
	return (2 * shared) / (left.length + right.length);
}

/** Rows that can be compared cell by cell: same cells, no nested tables. */
function comparable(a: Row, b: Row): boolean {
	return (
		a.length === b.length &&
		a.every(
			(cell, index) =>
				cell.gridSpan === b[index]!.gridSpan &&
				!cell.nestedTables?.length &&
				!b[index]!.nestedTables?.length &&
				!cell.sourceBlocks &&
				!b[index]!.sourceBlocks,
		)
	);
}

type RowStep =
	| { kind: 'pair'; a: number; b: number }
	| { kind: 'delete'; a: number }
	| { kind: 'insert'; b: number };

/** Pairs the rows of a changed stretch in order, maximising their total similarity. */
function pairGap(original: Row[], revised: Row[], aRows: number[], bRows: number[]): RowStep[] {
	const n = aRows.length;
	const m = bRows.length;
	const score = Array.from({ length: n + 1 }, () => new Float64Array(m + 1));
	for (let i = n - 1; i >= 0; i--)
		for (let j = m - 1; j >= 0; j--) {
			const a = original[aRows[i]!]!;
			const b = revised[bRows[j]!]!;
			const value = similarity(a, b);
			const pair = comparable(a, b) && value >= ROW_SIMILARITY ? value + score[i + 1]![j + 1]! : -1;
			score[i]![j] = Math.max(pair, score[i + 1]![j]!, score[i]![j + 1]!);
		}
	const steps: RowStep[] = [];
	let i = 0;
	let j = 0;
	while (i < n && j < m) {
		const a = original[aRows[i]!]!;
		const b = revised[bRows[j]!]!;
		const value = similarity(a, b);
		if (
			comparable(a, b) &&
			value >= ROW_SIMILARITY &&
			score[i]![j] === value + score[i + 1]![j + 1]!
		) {
			steps.push({ kind: 'pair', a: aRows[i++]!, b: bRows[j++]! });
		} else if (score[i]![j] === score[i + 1]![j]) steps.push({ kind: 'delete', a: aRows[i++]! });
		else steps.push({ kind: 'insert', b: bRows[j++]! });
	}
	while (i < n) steps.push({ kind: 'delete', a: aRows[i++]! });
	while (j < m) steps.push({ kind: 'insert', b: bRows[j++]! });
	return steps;
}

function rowSteps(original: Table, revised: Table, context: CompareContext): RowStep[] {
	const keys = (table: Table) => table.rows.map((row) => rowKey(row, context.granularity));
	const steps: RowStep[] = [];
	let aGap: number[] = [];
	let bGap: number[] = [];
	const flush = () => {
		steps.push(...pairGap(original.rows, revised.rows, aGap, bGap));
		aGap = [];
		bGap = [];
	};
	for (const op of diffKeys(keys(original), keys(revised))) {
		if (op.kind === 'delete') aGap.push(op.aIndex);
		else if (op.kind === 'insert') bGap.push(op.bIndex);
		else {
			flush();
			// Same text, but compared cell by cell for formatting when the cells line up.
			if (comparable(original.rows[op.aIndex]!, revised.rows[op.bIndex]!))
				steps.push({ kind: 'pair', a: op.aIndex, b: op.bIndex });
			else steps.push({ kind: 'delete', a: op.aIndex }, { kind: 'insert', b: op.bIndex });
		}
	}
	flush();
	return steps;
}

/** A result row's properties; a row of the original keeps no content control. */
function rowProperties(
	properties: TableRowProperties | undefined,
	revision?: TableRowProperties['revision'],
	original = false,
): TableRowProperties {
	const result: TableRowProperties = { ...properties };
	delete result.revision;
	if (original) delete result.contentControl;
	if (revision) result.revision = revision;
	return result;
}

function paragraphsOf(blocks: Block[]): Paragraph[] {
	return blocks.filter((block): block is Paragraph => block.type === 'paragraph');
}

function deletedRow(row: Row, context: CompareContext): Row {
	return row.map((cell) => ({
		...context.importCell(cell),
		paragraphs: paragraphsOf(wholeStory(cell.paragraphs, 'delete', context)),
	}));
}
function insertedRow(row: Row, context: CompareContext): Row {
	return row.map((cell) => ({
		...cell,
		paragraphs: paragraphsOf(wholeStory(cell.paragraphs, 'insert', context)),
	}));
}

/** An original table shown deleted: every row is a tracked row deletion. */
export function deletedTable(table: Table, context: CompareContext): Table {
	const result: Table = {
		...table,
		id: context.blockId(),
		rows: table.rows.map((row) => deletedRow(row, context)),
		rowProperties: table.rows.map((_, index) =>
			rowProperties(table.rowProperties?.[index], context.revision('delete'), true),
		),
	};
	delete result.contentControl;
	const style = context.tableStyle(table.style);
	if (style) result.style = style;
	else delete result.style;
	return result;
}

/** A revised table shown inserted: every row is a tracked row insertion. */
export function insertedTable(table: Table, context: CompareContext): Table {
	return {
		...table,
		rows: table.rows.map((row) => insertedRow(row, context)),
		rowProperties: table.rows.map((_, index) =>
			rowProperties(table.rowProperties?.[index], context.revision('insert')),
		),
	};
}

/**
 * Compares two tables row by row: rows with the same text pair, rows of a changed stretch pair
 * by their words, paired rows compare cell by cell and the rest are tracked row insertions and
 * deletions. A revised table whose rows the writer cannot restructure is shown replaced.
 */
export function compareTables(original: Table, revised: Table, context: CompareContext): Block[] {
	const steps = rowSteps(original, revised, context);
	const structural = steps.some((step) => step.kind !== 'pair');
	if (structural && !revised.structureEditable) {
		context.warn(
			'A table whose rows cannot be restructured is shown deleted and inserted whole instead of row by row.',
		);
		return [deletedTable(original, context), insertedTable(revised, context)];
	}
	const rows: Row[] = [];
	const properties: TableRowProperties[] = [];
	for (const step of steps) {
		if (step.kind === 'pair') {
			const a = original.rows[step.a]!;
			rows.push(
				revised.rows[step.b]!.map((cell, index) => ({
					...cell,
					paragraphs: paragraphsOf(compareStory(a[index]!.paragraphs, cell.paragraphs, context)),
				})),
			);
			properties.push(rowProperties(revised.rowProperties?.[step.b]));
		} else if (step.kind === 'delete') {
			rows.push(deletedRow(original.rows[step.a]!, context));
			properties.push(
				rowProperties(original.rowProperties?.[step.a], context.revision('delete'), true),
			);
		} else {
			rows.push(insertedRow(revised.rows[step.b]!, context));
			properties.push(rowProperties(revised.rowProperties?.[step.b], context.revision('insert')));
		}
	}
	const result: Table = { ...revised, rows };
	if (revised.rowProperties || properties.some((entry) => Object.keys(entry).length))
		result.rowProperties = properties;
	return [result];
}
