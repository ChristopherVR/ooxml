import { cellContentBottom } from './table-content';
import type { LayoutParagraph, LayoutTable } from './input';
import type { LayoutFootnoteBox } from './result';
import type { RowLayout } from './table-layout';
import type { PageCursor } from './page-cursor';
import type { ReferenceFootnotes } from './reference-footnotes';
import { cellParagraphs } from './table-cell-layout';
interface RowReference {
	note: LayoutFootnoteBox;
	bottom: number;
}
/** Preview only the bounded visible prefix, avoiding repeated copies of a long row's tail. */
function prefix(row: RowLayout, cut: number, cellCuts?: readonly number[]): RowLayout {
	const cells = row.cells.map((cell, index) => {
		const contentCut = (cellCuts?.[index] ?? cut) - (row.cellPadding?.[index]?.bottom ?? 0);
		const boxes = [];
		for (const box of cell) {
			if (box.yPx >= contentCut + 1e-6) break;
			if (box.yPx + box.heightPx + (box.tableFrameInsets?.bottom ?? 0) <= contentCut + 1e-6) {
				boxes.push(box);
				continue;
			}
			const lines = [];
			for (const line of box.lines) {
				if (
					box.yPx + line.yPx + line.heightPx + (box.tableFrameInsets?.bottom ?? 0) >
					contentCut + 1e-6
				)
					break;
				lines.push(line);
			}
			if (lines.length)
				boxes.push({ ...box, lines, heightPx: lines.at(-1)!.yPx + lines.at(-1)!.heightPx });
			// A nested row can have another paragraph beginning at the same y in its other cell.
			continue;
		}
		return boxes;
	});
	if (
		cells.every(
			(cell, index) =>
				cell.length === row.cells[index]!.length &&
				cell.every(
					(box, boxIndex) => box.lines.length === row.cells[index]![boxIndex]!.lines.length,
				),
		)
	)
		return row;
	return {
		...row,
		cells,
		heightPx: Math.max(
			0,
			...cells.map((cell, index) => cellContentBottom(cell, row.cellPadding?.[index]?.bottom ?? 0)),
		),
	};
}
export class TableFootnotes {
	private readonly paragraphs = new Map<string, LayoutParagraph>();
	private readonly placed = new Set<string>();
	private readonly fragment = new Map<string, RowReference>();
	constructor(
		private readonly cursor: PageCursor,
		private readonly table: LayoutTable,
		private readonly height: () => number,
		private readonly resolve: ReferenceFootnotes,
	) {
		for (const row of table.rows)
			for (const cell of row.cells)
				for (const paragraph of cellParagraphs(cell))
					if (paragraph.footnotes?.length) this.paragraphs.set(paragraph.id, paragraph);
	}
	private references(row: RowLayout): RowReference[] {
		const refs = new Map<string, RowReference>();
		for (const cell of row.cells)
			for (const box of cell) {
				const source = this.paragraphs.get(box.blockId);
				if (!source) continue;
				for (const note of this.resolve(
					source,
					box.lines,
					(box.lines[0]?.sourceStart ?? 0) === 0,
				)) {
					if (this.placed.has(note.id)) continue;
					const line =
						box.lines.find((line) =>
							line.fragments.some(
								(fragment) => source.runs[fragment.runIndex]?.footnoteReferenceId === note.id,
							),
						) ?? box.lines[0];
					const bottom = Math.min(
						row.heightPx,
						box.yPx +
							(line?.yPx ?? 0) +
							(line?.heightPx ?? box.heightPx) +
							(box.tableFrameInsets?.bottom ?? 0),
					);
					const prior = refs.get(note.id);
					if (!prior || bottom < prior.bottom) refs.set(note.id, { note, bottom });
				}
			}
		return [...refs.values()];
	}
	private hold(refs: Map<string, RowReference>, minimumBody = 0): void {
		const referenceBottom = Math.max(0, ...[...refs.values()].map((ref) => ref.bottom));
		const bodyHeight = Math.max(referenceBottom, minimumBody);
		this.cursor.holdFootnotes(
			this.table.id,
			[...refs.values()].map((ref) => ref.note),
			bodyHeight,
			this.cursor.usedColumnHeightPx + referenceBottom,
		);
	}
	private holdCandidate(row: RowLayout): void {
		const refs = new Map(this.fragment);
		for (const ref of this.references(row))
			if (!refs.has(ref.note.id))
				refs.set(ref.note.id, { ...ref, bottom: this.height() + ref.bottom });
		const referenceBottom = Math.max(0, ...[...refs.values()].map((ref) => ref.bottom));
		const contentBottom =
			this.height() + Math.max(0, ...row.cells.map((cell) => cellContentBottom(cell)));
		const paddingBottom = Math.max(0, ...(row.cellPadding ?? []).map((padding) => padding.bottom));
		this.hold(
			refs,
			contentBottom <= referenceBottom + 1e-6
				? this.height() + row.heightPx
				: referenceBottom + paddingBottom,
		);
	}
	/** Largest complete prefix that keeps each new marker with available note content. */
	choose(
		row: RowLayout,
		atomic: boolean,
	): { layout: RowLayout; cut: number; cellCuts?: number[] } | undefined {
		this.cursor.holdFootnotes(this.table.id, [], 0);
		const available = this.cursor.remainingHeightPx() - this.height();
		if (row.heightPx <= available + 1e-6) {
			this.holdCandidate(row);
			if (row.heightPx <= this.cursor.remainingHeightPx() - this.height() + 1e-6)
				return { layout: row, cut: row.heightPx };
		}
		if (!atomic) {
			const cuts = new Set<number>();
			for (const [index, cell] of row.cells.entries())
				for (const box of cell) {
					if (box.yPx > available + 1e-6) break;
					for (const line of box.lines) {
						const bottom =
							box.yPx +
							line.yPx +
							line.heightPx +
							(box.tableFrameInsets?.bottom ?? 0) +
							(row.cellPadding?.[index]?.bottom ?? 0);
						if (bottom > available + 1e-6) break;
						if (bottom > 0) cuts.add(bottom);
					}
				}
			for (const cut of [...cuts].sort((a, b) => b - a)) {
				const candidate = prefix(row, cut);
				this.holdCandidate(candidate);
				if (candidate.heightPx <= this.cursor.remainingHeightPx() - this.height() + 1e-6) {
					// A delayed marker constrains its own cell; neighboring cells without references
					// can consume the remaining row space (native padded split-row behavior).
					const space = this.cursor.remainingHeightPx() - this.height();
					const cellCuts = row.cells.map((cell) =>
						cell.some((box) => this.paragraphs.has(box.blockId)) ? cut : space,
					);
					const expanded = prefix(row, cut, cellCuts);
					this.holdCandidate(expanded);
					if (expanded.heightPx <= this.cursor.remainingHeightPx() - this.height() + 1e-6)
						return { layout: expanded, cut, cellCuts };
					this.holdCandidate(candidate);
					return { layout: candidate, cut };
				}
			}
		}
		this.holdFragment();
		return undefined;
	}
	add(row: RowLayout, repeated: boolean): void {
		if (repeated) return;
		for (const ref of this.references(row))
			if (!this.fragment.has(ref.note.id))
				this.fragment.set(ref.note.id, { ...ref, bottom: this.height() + ref.bottom });
	}
	holdFragment(): void {
		this.hold(this.fragment);
	}
	flushed(): void {
		for (const id of this.fragment.keys()) this.placed.add(id);
		this.fragment.clear();
	}
}
