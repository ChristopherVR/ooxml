// Repeating regions of the headless merge: the content between `MERGEFIELD TableStart:<name>` and
// `MERGEFIELD TableEnd:<name>` repeats once per row of `regions[<name>]`. Markers in the rows of
// one table repeat those rows; markers in the paragraphs of one story repeat the blocks between
// them. Copies get new ids, and the writer is told which source block or row each copies, so a
// copy keeps every property and XML child of its source.
import { rowEditKey, type BlockCopy } from '../block-copies';
import { documentParagraphs, mapStoryBlocks } from '../cell-blocks';
import { fieldResultSpans, type FieldResultSpan } from '../field-result-spans';
import type { Block, DocumentModel, Paragraph, Table, TableCell, TextRun } from '../model';
import type { MailMergeRecord } from './data-source';
import { mergeFieldColumn } from './field-values';
import { forEachBlockList, mapBlockLists, type MergeStory } from './stories';

export interface MergeRegionInfo {
	name: string;
	/** `rows`: the markers sit in the rows of one table; `paragraphs`: in one story's paragraphs. */
	kind: 'rows' | 'paragraphs';
	story: MergeStory;
}

export interface MergeRegionResult extends MergeRegionInfo {
	/** How many times the region was written: the number of rows given for it. */
	rows: number;
}

/** A region marker field: `TableStart:<name>` or `TableEnd:<name>`. */
export function regionMarker(instruction: string): { start: boolean; name: string } | undefined {
	const column = mergeFieldColumn(instruction);
	const match = column && /^Table(Start|End):(.+)$/i.exec(column);
	return match ? { start: match[1]!.toLowerCase() === 'start', name: match[2]! } : undefined;
}

interface Marker {
	start: boolean;
	name: string;
}

const markersOf = (paragraph: Paragraph): Marker[] =>
	fieldResultSpans(paragraph.runs, { resultless: true }).flatMap((span) => {
		const marker = regionMarker(span.instr);
		return marker ? [marker] : [];
	});

/** Pairs markers in order; `at` is the position (row or block index) of each marker. */
function pairMarkers(markers: { marker: Marker; at: number }[]) {
	const pairs: { name: string; from: number; to: number }[] = [];
	let open: { name: string; at: number } | undefined;
	for (const { marker, at } of markers) {
		if (marker.start) {
			// A region inside another is not supported: the inner markers are only removed.
			if (!open) open = { name: marker.name, at };
		} else if (open && marker.name === open.name) {
			pairs.push({ name: open.name, from: open.at, to: at });
			open = undefined;
		}
	}
	return pairs;
}

/** Each region's info, in story order. */
export function listMergeRegions(model: DocumentModel): MergeRegionInfo[] {
	const regions: MergeRegionInfo[] = [];
	forEachBlockList(model, (blocks, list) => {
		if (list.cell) return;
		for (const block of blocks)
			if (block.type === 'table')
				for (const pair of pairMarkers(rowMarkers(block)))
					regions.push({ name: pair.name, kind: 'rows', story: list.story });
		for (const pair of pairMarkers(blockMarkers(blocks)))
			regions.push({ name: pair.name, kind: 'paragraphs', story: list.story });
	});
	return regions;
}

const rowMarkers = (table: Table) =>
	table.rows.flatMap((row, at) =>
		row.flatMap((cell) => cell.paragraphs.flatMap(markersOf).map((marker) => ({ marker, at }))),
	);

const blockMarkers = (blocks: readonly Block[]) =>
	blocks.flatMap((block, at) =>
		block.type === 'paragraph' ? markersOf(block).map((marker) => ({ marker, at })) : [],
	);

export interface ExpandedRegions {
	model: DocumentModel;
	/** The record each copied paragraph's fields read. */
	scope: Map<string, MailMergeRecord>;
	copies: Map<string, BlockCopy>;
	regions: MergeRegionResult[];
}

function idSuffix(model: DocumentModel): string {
	const ids: string[] = [];
	forEachBlockList(model, (blocks) => {
		for (const block of blocks) ids.push(block.id);
	});
	let suffix = '-mm';
	while (ids.some((id) => id.includes(suffix))) suffix += 'm';
	return suffix;
}

/** Repeats every region of `model` for its rows; the markers themselves are left in place. */
export function expandRegions(
	model: DocumentModel,
	data: Readonly<Record<string, readonly MailMergeRecord[]>>,
): ExpandedRegions {
	const scope = new Map<string, MailMergeRecord>();
	const copies = new Map<string, BlockCopy>();
	const regions: MergeRegionResult[] = [];
	const suffix = idSuffix(model);
	let serial = 0;
	const copyKey = () => {
		const tag = `${suffix}${serial++}`;
		return (id: string) => `${id}${tag}`;
	};
	const rowsFor = (name: string) => (Object.hasOwn(data, name) ? (data[name] ?? []) : []);
	const scopeAll = (blocks: readonly Block[], record: MailMergeRecord) => {
		for (const paragraph of documentParagraphs(blocks))
			if (!scope.has(paragraph.id)) scope.set(paragraph.id, record);
	};
	const expanded = mapBlockLists(model, (blocks, list) => {
		if (list.cell) return blocks;
		let changed = false;
		let next = blocks.map((block): Block => {
			if (block.type !== 'table') return block;
			const pairs = pairMarkers(rowMarkers(block));
			if (!pairs.length) return block;
			changed = true;
			copies.set(rowEditKey(block.id), { sourceId: '', rekey: (id) => id });
			return repeatRows(block, pairs, list.story);
		});
		const pairs = pairMarkers(blockMarkers(next));
		for (const pair of pairs.reverse()) {
			changed = true;
			const source = next.slice(pair.from, pair.to + 1);
			const records = rowsFor(pair.name);
			regions.push({
				name: pair.name,
				kind: 'paragraphs',
				story: list.story,
				rows: records.length,
			});
			const output = records.flatMap((record, index) => {
				const copy = index ? copyBlocks(source, list.story === 'textBox') : source;
				scopeAll(copy, record);
				return copy;
			});
			next = [...next.slice(0, pair.from), ...output, ...next.slice(pair.to + 1)];
		}
		next = next.filter((block) => block.type !== 'table' || block.rows.length);
		if (!next.length) {
			const first = blocks.find((block): block is Paragraph => block.type === 'paragraph');
			next = [{ type: 'paragraph', id: first?.id ?? `empty${suffix}`, runs: [{ text: '' }] }];
		}
		return changed ? next : blocks;
	});

	function copyBlocks(source: Block[], textBox: boolean): Block[] {
		const rekey = copyKey();
		return source.map((block) => {
			const [copy] = mapStoryBlocks(
				[block],
				(paragraph) => ({ ...paragraph, id: rekey(paragraph.id) }),
				rekey,
			);
			const id = rekey(block.id);
			if (!textBox) copies.set(id, { sourceId: block.id, rekey });
			return { ...copy!, id };
		});
	}

	function repeatRows(
		table: Table,
		pairs: { name: string; from: number; to: number }[],
		story: MergeStory,
	): Table {
		const rows: TableCell[][] = [];
		const properties: NonNullable<Table['rowProperties']> = [];
		const hasProperties = Boolean(table.rowProperties);
		let at = 0;
		const keep = (index: number) => {
			rows.push(table.rows[index]!);
			if (hasProperties) properties.push(table.rowProperties![index] ?? {});
		};
		for (const pair of pairs) {
			while (at < pair.from) keep(at++);
			const records = rowsFor(pair.name);
			regions.push({ name: pair.name, kind: 'rows', story, rows: records.length });
			records.forEach((record, index) => {
				const rekey = index ? copyKey() : (id: string) => id;
				for (let row = pair.from; row <= pair.to; row++) {
					const copy = table.rows[row]!.map((cell) => ({
						...cell,
						paragraphs: cell.paragraphs.map((paragraph) => {
							const id = rekey(paragraph.id);
							if (index) copies.set(id, { sourceId: paragraph.id, rekey });
							scope.set(id, record);
							return { ...paragraph, id };
						}),
					}));
					rows.push(copy);
					if (hasProperties) properties.push(table.rowProperties![row] ?? {});
				}
			});
			at = pair.to + 1;
		}
		while (at < table.rows.length) keep(at++);
		return { ...table, rows, ...(hasProperties && { rowProperties: properties }) };
	}

	return { model: expanded, scope, copies, regions };
}

/** The runs without the fields of `span`s that `remove` picks; the paragraph itself if none. */
export function withoutFields(
	paragraph: Paragraph,
	remove: (span: FieldResultSpan) => boolean,
): Paragraph {
	const spans = fieldResultSpans(paragraph.runs, { resultless: true }).filter(remove);
	if (!spans.length) return paragraph;
	const dropped = new Set<number>();
	for (const span of spans) for (let at = span.from; at <= span.to; at++) dropped.add(at);
	const runs: TextRun[] = paragraph.runs.filter((_, index) => !dropped.has(index));
	return { ...paragraph, runs: runs.length ? runs : [{ text: '' }] };
}

/** Removes every region marker field. */
export const withoutRegionMarkers = (model: DocumentModel): DocumentModel =>
	mapBlockLists(model, (blocks) => {
		let changed = false;
		const next = blocks.map((block) => {
			if (block.type !== 'paragraph') return block;
			const paragraph = withoutFields(block, (span) => regionMarker(span.instr) !== undefined);
			changed ||= paragraph !== block;
			return paragraph;
		});
		return changed ? next : blocks;
	});
