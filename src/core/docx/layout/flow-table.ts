import { independentNoteStories, measuredNoteReferences } from './table-note-stories';
import { layoutBorderedRow, TableFragmentBorders } from './table-border-layout';
import { nestedFrameProfile } from './nested-frame-profile';
import { cellParagraphs } from './table-cell-layout';
import { TableFootnotes } from './table-footnotes';
import type { ReferenceFootnotes } from './reference-footnotes';
import type { PageCursor } from './page-cursor';
import { splitRowAtHeight, type RowLayout } from './table-layout';
import type { TextMeasurer } from './measure';
import type { LayoutTable, LayoutTableRow } from './input';
import type { LayoutTableBox, LayoutTableRowBox } from './result';

/**
 * Places a table row by row, splitting a row across a page/column boundary
 * unless it is `cantSplit`, and repeating `tblHeader` rows at the top of
 * every page/column fragment the table continues onto. Each fragment the
 * table touches becomes its own `LayoutTableBox` sharing the table's
 * `blockId`, since a single result box can only live in one column.
 */
export function placeTable(
	cursor: PageCursor,
	table: LayoutTable,
	measurer: TextMeasurer,
	note: (message: string) => void,
	resolveFootnotes?: ReferenceFootnotes,
): void {
	// Word only repeats the contiguous header group starting at the first row.
	// A tblHeader flag after the first body row is ignored (ECMA-376 17.4.50).
	const firstBodyRow = table.rows.findIndex((row) => !row.isHeader);
	const headerRows = table.rows.slice(0, firstBodyRow < 0 ? table.rows.length : firstBodyRow);

	// Table indent, or alignment of the grid within the column.
	const xPosition = () => {
		const width =
			(table.widthPx ?? cursor.columnWidthPx) +
			(activeBorders ? (profile!.leftPx + profile!.rightPx) / 2 : 0);
		return table.alignment === 'center'
			? Math.max(0, (cursor.columnWidthPx - width) / 2)
			: table.alignment === 'right'
				? Math.max(0, cursor.columnWidthPx - width)
				: (table.indentPx ?? 0);
	};

	const profile = nestedFrameProfile(table, table.widthPx ?? cursor.columnWidthPx);
	const activeBorders =
		profile !== undefined && (profile.leftPx > 0 || profile.rightPx > 0 || profile.bottomPx > 0);
	if (!profile && table.rows.some((row) => row.cells.some((cell) => cell.borders)))
		note(
			'Ordinary table border occupancy remains approximate for missing edges, non-solid or conflicting shared borders and changing row grids.',
		);

	const layoutRow = (
		row: LayoutTableRow,
		width: number,
		m: TextMeasurer,
		n: (message: string) => void,
		remaining?: import('./result').LayoutParagraphBox[][],
	) =>
		layoutBorderedRow(
			row,
			activeBorders ? (table.widthPx ?? width) : width,
			m,
			n,
			profile,
			remaining,
		);
	const fragmentBorders = new TableFragmentBorders();
	let fragmentRows: LayoutTableRowBox[] = [];
	let fragmentHeight = 0;
	const footnotes = resolveFootnotes
		? new TableFootnotes(cursor, table, () => fragmentHeight, resolveFootnotes)
		: undefined;

	const toRowBox = (layout: RowLayout, repeated: boolean): LayoutTableRowBox => {
		fragmentHeight += fragmentBorders.remove(fragmentRows.at(-1));
		footnotes?.add(layout, repeated);
		const box: LayoutTableRowBox = {
			yPx: fragmentHeight,
			heightPx: layout.heightPx,
			repeated,
			cells: layout.cells,
			geometry: layout.geometry,
			...(layout.cellTables ? { cellTables: layout.cellTables } : {}),
		};
		fragmentHeight += layout.heightPx;
		fragmentBorders.record(layout.borderBottomPx);
		return box;
	};

	function flushFragment() {
		fragmentHeight += fragmentBorders.restore(fragmentRows.at(-1));
		if (!fragmentRows.length) {
			footnotes?.holdFragment();
			return;
		}
		const box: LayoutTableBox = {
			kind: 'table',
			blockId: table.id,
			...(activeBorders
				? {
						widthPx:
							(table.widthPx ?? cursor.columnWidthPx) + (profile!.leftPx + profile!.rightPx) / 2,
					}
				: {}),
			...(xPosition() ? { xPx: xPosition() } : {}),
			yPx: 0,
			heightPx: fragmentHeight,
			rows: fragmentRows,
		};
		footnotes?.holdFragment();
		cursor.place(box, fragmentHeight);
		footnotes?.flushed();
		fragmentRows = [];
		fragmentHeight = 0;
	}
	function startNewFragment() {
		flushFragment();
		cursor.newColumn();
		addHeaders();
	}
	function addHeaders() {
		for (const header of headerRows)
			fragmentRows.push(toRowBox(layoutRow(header, cursor.columnWidthPx, measurer, note), true));
	}

	function placeRow(row: LayoutTableRow, initial: RowLayout, cantSplit: boolean) {
		let remainingRow: RowLayout | null = initial;
		while (remainingRow) {
			fragmentHeight += fragmentBorders.remove(fragmentRows.at(-1));
			const minimum = cantSplit
				? Math.min(remainingRow.heightPx, cursor.columnHeightPx)
				: Math.min(
						...remainingRow.cells.flatMap((cell, index) =>
							cell.map(
								(paragraph) =>
									paragraph.yPx +
									(paragraph.lines[0]?.yPx ?? 0) +
									(paragraph.lines[0]?.heightPx ?? 0) +
									(remainingRow!.borderBottomPx !== undefined
										? (remainingRow!.cellPadding?.[index]?.bottom ?? 0) +
											(paragraph.tableFrameInsets?.bottom ?? 0)
										: 0),
							),
						),
					);
			if (fragmentRows.every((item) => item.repeated)) {
				const width = cursor.columnWidthPx;
				cursor.advancePastFootnotes(fragmentHeight + (Number.isFinite(minimum) ? minimum : 0));
				if (width !== cursor.columnWidthPx) {
					remainingRow = layoutRow(
						remainingRow === initial ? row : { ...row, heightPx: 0, heightRule: 'atLeast' },
						cursor.columnWidthPx,
						measurer,
						note,
						remainingRow.cells,
					);
					if (fragmentRows.length) {
						fragmentRows = [];
						fragmentHeight = 0;
						addHeaders();
					}
					continue;
				}
			}
			if (footnotes) {
				const selected = footnotes.choose(remainingRow, cantSplit);
				if (selected) {
					if (selected.layout === remainingRow) {
						fragmentRows.push(toRowBox(remainingRow, false));
						remainingRow = null;
						continue;
					}
					const { before, after } = splitRowAtHeight(
						remainingRow,
						selected.cut + 1e-6,
						selected.cellCuts?.map((cut) => cut + 1e-6),
						selected.paragraphCuts
							? new Map([...selected.paragraphCuts].map(([id, cut]) => [id, cut + 1e-6]))
							: undefined,
					);
					fragmentRows.push(toRowBox(before, false));
					const previousWidth = cursor.columnWidthPx;
					startNewFragment();
					// Independent inner stories resume at different tokens; source reflow restores each
					// cell origin and places following direct text after the longest continuation.
					remainingRow =
						after && (selected.paragraphCuts || previousWidth !== cursor.columnWidthPx)
							? layoutRow(
									{ ...row, heightPx: 0, heightRule: 'atLeast' },
									cursor.columnWidthPx,
									measurer,
									note,
									after.cells,
								)
							: after;
					continue;
				}
				if (fragmentRows.length || !cursor.atColumnTop) {
					const previousWidth = cursor.columnWidthPx;
					startNewFragment();
					if (previousWidth !== cursor.columnWidthPx)
						remainingRow = layoutRow(
							remainingRow === initial ? row : { ...row, heightPx: 0, heightRule: 'atLeast' },
							cursor.columnWidthPx,
							measurer,
							note,
							remainingRow === initial ? undefined : remainingRow.cells,
						);
					continue;
				}
			}
			const remaining = cursor.remainingHeightPx() - fragmentHeight;
			if (remainingRow.heightPx <= remaining) {
				fragmentRows.push(toRowBox(remainingRow, false));
				remainingRow = null;
				continue;
			}
			if (cantSplit) {
				if (fragmentRows.length || !cursor.atColumnTop) {
					startNewFragment();
					remainingRow = layoutRow(row, cursor.columnWidthPx, measurer, note);
				}
				const noteWidth = cursor.columnWidthPx;
				cursor.advancePastFootnotes(
					fragmentHeight + Math.min(remainingRow.heightPx, cursor.columnHeightPx),
				);
				if (noteWidth !== cursor.columnWidthPx) {
					remainingRow = layoutRow(row, cursor.columnWidthPx, measurer, note);
					if (fragmentRows.length) {
						fragmentRows = [];
						fragmentHeight = 0;
						addHeaders();
					}
				}
				const freshRemaining = cursor.remainingHeightPx() - fragmentHeight;
				if (remainingRow.heightPx > freshRemaining)
					note('A table row marked "keep row together" is taller than an empty page/column.');
				fragmentRows.push(toRowBox(remainingRow, false));
				remainingRow = null;
				continue;
			}
			const { before, after } = splitRowAtHeight(remainingRow, Math.max(0, remaining));
			if (
				!before.cells.some((cell) => cell.length) &&
				cursor.atColumnTop &&
				fragmentRows.every((row) => row.repeated)
			) {
				note(
					'A table row cannot fit a line in an empty page/column and was placed without splitting.',
				);
				fragmentRows.push(toRowBox(remainingRow, false));
				remainingRow = null;
				continue;
			}
			const previousWidth = cursor.columnWidthPx;
			if (before.cells.some((cell) => cell.length)) fragmentRows.push(toRowBox(before, false));
			startNewFragment();
			remainingRow =
				after && previousWidth !== cursor.columnWidthPx
					? layoutRow(
							{ ...row, heightPx: 0, heightRule: 'atLeast' },
							cursor.columnWidthPx,
							measurer,
							note,
							after.cells,
						)
					: after;
		}
	}

	for (const row of table.rows) {
		const layout = layoutRow(row, cursor.columnWidthPx, measurer, note);
		if (
			activeBorders &&
			layout.heightPx > cursor.remainingHeightPx() &&
			row.cells.some(
				(cell) =>
					(!measuredNoteReferences(row.cells.flatMap(cellParagraphs)) ||
						!independentNoteStories(cell)) &&
					cell.blocks?.some((block) => block.kind === 'table') &&
					cellParagraphs(cell).some((p) => p.footnotes?.length),
			)
		)
			note(
				'Independent inner-cell continuation around nested note references remains approximate when an enclosing bordered row splits.',
			);
		placeRow(
			row,
			layout,
			Boolean(row.cantSplit) ||
				headerRows.includes(row) ||
				(!!footnotes && row.heightRule === 'exact'),
		);
	}
	flushFragment();
}
