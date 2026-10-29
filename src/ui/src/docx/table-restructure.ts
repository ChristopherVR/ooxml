// A merged or nested table in the editor holds one node per visible cell (a vertical merge is one
// node with a rowspan); the cells it continues through are only named, in `continueKeys`. This
// module lays such a node out on its grid and converts it to and from the core table model, so
// the core structure edits (ooxml-core/docx table-structure) run on it and their result comes
// back as editor nodes that keep the content of every surviving cell.
import type { Table, TableCell } from 'ooxml-core/docx';
import { rowSpanOf, continuationsOf, tableGrid } from 'ooxml-core/docx/ui';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import { schema } from './schema';

/** One position of a row: a cell node, or a cell the vertical merge above it continues through. */
export interface Slot {
	row: number;
	col: number;
	span: number;
	/** The cell node (the merge start, for a continuation). */
	cell: ProseMirrorNode;
	/** 0 for the cell itself, k for its k-th continuation row. */
	offset: number;
	/** First paragraph id of the source cell this slot stands for. */
	key: string;
}

export const keysOf = (cell: ProseMirrorNode): string[] => {
	try {
		const parsed: unknown = JSON.parse(String(cell.attrs.continueKeys ?? '[]'));
		return Array.isArray(parsed) ? parsed.map(String) : [];
	} catch {
		return [];
	}
};

export const gridOf = (table: ProseMirrorNode): number[] | undefined => {
	try {
		const parsed: unknown = table.attrs.grid ? JSON.parse(String(table.attrs.grid)) : undefined;
		return Array.isArray(parsed) ? parsed.map(Number) : undefined;
	} catch {
		return undefined;
	}
};

/** The slots of each row of a table node, left to right. */
export function slotsOf(table: ProseMirrorNode): Slot[][] {
	const pending = new Map<number, Slot[]>();
	const rows: Slot[][] = [];
	table.forEach((row, _offset, r) => {
		const carried = [...(pending.get(r) ?? [])].sort((a, b) => a.col - b.col);
		const slots: Slot[] = [];
		const cells: ProseMirrorNode[] = [];
		row.forEach((cell) => cells.push(cell));
		let col = 0;
		let next = 0;
		while (carried.length || next < cells.length) {
			const continued = carried[0];
			if (continued && continued.col === col) {
				slots.push(continued);
				col += continued.span;
				carried.shift();
				continue;
			}
			const cell = cells[next++];
			if (!cell) break;
			const span = Number(cell.attrs.colspan) || 1;
			const key = String(cell.attrs.sourceCellKey || '');
			slots.push({ row: r, col, span, cell, offset: 0, key });
			const rowspan = Number(cell.attrs.rowspan) || 1;
			const names = keysOf(cell);
			for (let k = 1; k < rowspan; k++)
				pending.set(r + k, [
					...(pending.get(r + k) ?? []),
					{ row: r + k, col, span, cell, offset: k, key: names[k - 1] ?? `${key}~${k}` },
				]);
			col += span;
		}
		rows.push(slots);
	});
	return rows;
}

const json = <T>(value: unknown): T | undefined => {
	try {
		return typeof value === 'string' ? (JSON.parse(value) as T) : undefined;
	} catch {
		return undefined;
	}
};

/** The core model of a table node: cell ids, spans, merges, widths and the formatting new cells copy. */
export function skeletonOf(table: ProseMirrorNode): Table {
	const rows = slotsOf(table).map((slots) =>
		slots.map((slot): TableCell => {
			const { cell } = slot;
			const base: TableCell =
				slot.offset === 0
					? { paragraphs: [] }
					: { paragraphs: [{ type: 'paragraph', id: slot.key, runs: [{ text: '' }] }] };
			if (slot.span > 1) base.gridSpan = slot.span;
			const width = cell.attrs.widthTwips as TableCell['widthTwips'] | null;
			if (width != null) base.widthTwips = width;
			if (slot.offset > 0) {
				base.verticalMerge = 'continue';
				return base;
			}
			if (Number(cell.attrs.rowspan) > 1) base.verticalMerge = 'restart';
			cell.forEach((child) => {
				if (child.type.name === 'paragraph')
					base.paragraphs.push({
						type: 'paragraph',
						id: String(child.attrs.id),
						runs: [{ text: child.textContent }],
					});
				else if (child.type.name === 'table')
					(base.sourceBlocks ??= []).push({ type: 'table', id: String(child.attrs.id), rows: [] });
				else (base.nestedTables ??= []).push({ rows: [] });
			});
			if (!base.paragraphs.length)
				base.paragraphs.push({ type: 'paragraph', id: slot.key, runs: [{ text: '' }] });
			if (cell.attrs.verticalAlign) base.verticalAlign = cell.attrs.verticalAlign;
			if (cell.attrs.directShadingFill) base.shadingFill = String(cell.attrs.directShadingFill);
			const borders = json<NonNullable<TableCell['borders']>>(cell.attrs.directBorders);
			if (borders) base.borders = borders;
			const margins = json<NonNullable<TableCell['margins']>>(cell.attrs.margins);
			if (margins) base.margins = margins;
			return base;
		}),
	);
	const model: Table = { type: 'table', id: String(table.attrs.id), rows };
	const grid = gridOf(table);
	if (grid) model.grid = grid as NonNullable<Table['grid']>;
	const properties: NonNullable<Table['rowProperties']> = [];
	table.forEach((row) => properties.push(structuredClone(row.attrs.properties ?? {})));
	if (properties.some((entry) => Object.keys(entry).length)) model.rowProperties = properties;
	return model;
}

const FORMAT_KEYS = [
	'style',
	'align',
	'justification',
	'direction',
	'spacingBeforeTwips',
	'spacingAfterTwips',
	'lineSpacingTwips',
	'lineSpacingRule',
	'indentLeftTwips',
	'indentRightTwips',
	'indentStartTwips',
	'indentEndTwips',
	'firstLineTwips',
	'hangingTwips',
] as const;

function paragraphNode(
	model: Table['rows'][number][number]['paragraphs'][number],
	known: Map<string, ProseMirrorNode>,
): ProseMirrorNode {
	const existing = known.get(model.id);
	if (existing) return existing;
	const attrs: Record<string, unknown> = { id: model.id };
	const source = model as unknown as Record<string, unknown>;
	for (const key of FORMAT_KEYS) if (source[key] !== undefined) attrs[key] = source[key];
	return schema.nodes.paragraph.create(attrs);
}

/** A table node for `result`, a structure edit of `old`'s skeleton; surviving cells keep their content. */
export function nodeFromResult(old: ProseMirrorNode, result: Table): ProseMirrorNode | undefined {
	const grid = tableGrid(result);
	if (!grid) return undefined;
	const cellOf = new Map<string, ProseMirrorNode>();
	const known = new Map<string, ProseMirrorNode>();
	for (const row of slotsOf(old))
		for (const slot of row) {
			if (slot.offset) continue;
			slot.cell.forEach((child) => {
				if (child.type.name !== 'paragraph') return;
				known.set(String(child.attrs.id), child);
				cellOf.set(String(child.attrs.id), slot.cell);
			});
		}
	const rows = grid.rows.map((entries, r) => {
		const cells: ProseMirrorNode[] = [];
		for (const entry of entries) {
			const model = entry.cell;
			if (model.verticalMerge === 'continue') continue;
			const rowspan = rowSpanOf(grid, entry);
			const keys = continuationsOf(grid, entry).map((c) => c.cell.paragraphs[0]?.id ?? '');
			const ids = model.paragraphs.map((p) => p.id);
			const source = ids.map((id) => cellOf.get(id)).find(Boolean);
			const attrs: Record<string, unknown> = {
				...(source?.attrs ?? {}),
				colspan: entry.span,
				rowspan,
				widthTwips: model.widthTwips ?? null,
				continueKeys: keys.length ? JSON.stringify(keys) : null,
			};
			if (!source) {
				Object.assign(attrs, {
					sourceCellKey: ids[0] ?? '',
					verticalAlign: model.verticalAlign ?? null,
					shadingFill: model.shadingFill ?? null,
					directShadingFill: model.shadingFill ?? null,
					borders: model.borders ? JSON.stringify(model.borders) : null,
					directBorders: model.borders ? JSON.stringify(model.borders) : null,
					margins: model.margins ? JSON.stringify(model.margins) : null,
					contentControl: null,
				});
			}
			const content: ProseMirrorNode[] = [];
			const wanted = new Set(ids);
			source?.forEach((child) => {
				if (child.type.name !== 'paragraph' || wanted.has(String(child.attrs.id)))
					content.push(child);
			});
			const present = new Set(content.map((child) => String(child.attrs.id)));
			for (const paragraph of model.paragraphs)
				if (!present.has(paragraph.id)) content.push(paragraphNode(paragraph, known));
			cells.push(schema.nodes.tableCell.create(attrs, content));
		}
		const properties = result.rowProperties?.[r];
		const height = properties?.heightTwips;
		return schema.nodes.tableRow.create(
			{
				heightPx: height ? height / 15 : null,
				properties: properties && Object.keys(properties).length ? properties : null,
			},
			cells,
		);
	});
	return schema.nodes.table.create(
		{
			...old.attrs,
			// A table that was edited this way is saved through the merge-aware writer from now on.
			structureEditable: false,
			restructurable: true,
			grid: result.grid ? JSON.stringify(result.grid) : null,
		},
		rows,
	);
}
