import { reflowNestedFragments } from './nested-fragment-edges';
import { tableParagraphKeeps } from './table-paragraph-keeps';
import { widowControlCut } from './table-widow-cuts';
import { independentNoteStories, measuredNoteReferences } from './table-note-stories';
import { layoutBorderedRow, TableFragmentBorders } from './table-border-layout';
import { nestedFrameProfile } from './nested-frame-profile';
import { cellParagraphs } from './table-cell-layout';
import { TableFootnotes } from './table-footnotes';
import type { ReferenceFootnotes } from './reference-footnotes';
import type { PageCursor } from './page-cursor';
import { MergedRowPlan } from './table-merged-rows';
import { splitRowAtHeight, type RowLayout } from './table-layout';
import { resolveTableAutofit } from './table-autofit';
import { placeFloatingTable } from './flow-floating-table';
import type { TextMeasurer } from './measure';
import type { LayoutTable, LayoutTableRow, OversizedCantSplitRow } from './input';
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
	source: LayoutTable,
	measurer: TextMeasurer,
	note: (message: string) => void,
	resolveFootnotes?: ReferenceFootnotes,
	oversizedCantSplitRow: OversizedCantSplitRow = 'split',
): void {
	// AutoFit columns are sized for the text column the table starts in.
	const table = resolveTableAutofit(source, cursor.columnWidthPx, measurer, note);
	if (
		table.floating &&
		placeFloatingTable(cursor, { ...table, floating: table.floating }, measurer, note)
	)
		return;
	// Word only repeats the contiguous header group starting at the first row.
	// A tblHeader flag after the first body row is ignored (ECMA-376 17.4.50).
	const firstBodyRow = table.rows.findIndex((row) => !row.isHeader);
	const headerRows = table.rows.slice(0, firstBodyRow < 0 ? table.rows.length : firstBodyRow);
	// Continuation rows retain their source cells even when their height constraint is reset.
	const nativeFragments = new Map(table.rows.map((row) => [row, reflowNestedFragments(row)]));
	const paragraphKeeps = tableParagraphKeeps(table);
	const hasParagraphKeeps = table.rows.some((row) =>
		row.cells.some((cell) => cellParagraphs(cell).some((paragraph) => paragraph.keepLines)),
	);
	let keepWarning = false;

	// Table indent, or alignment of the grid within the column.
	const xPosition = () => {
		const width =
			(table.widthPx ?? cursor.columnWidthPx) +
			(activeBorders ? (profile!.leftPx + profile!.rightPx) / 2 : 0);
		// A table wider than the text column still aligns: it overhangs the margins (#387).
		return table.alignment === 'center'
			? (cursor.columnWidthPx - width) / 2
			: table.alignment === 'right'
				? cursor.columnWidthPx - width
				: (table.indentPx ?? 0);
	};

	const profile = nestedFrameProfile(table, table.widthPx ?? cursor.columnWidthPx, true);
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
		sourceRow = row,
	) =>
		layoutBorderedRow(
			row,
			activeBorders ? (table.widthPx ?? width) : width,
			m,
			n,
			profile,
			remaining,
			nativeFragments.get(sourceRow) ?? false,
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
		alignMergedCells(fragmentRows);
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
	// Header rows repeat only once the table's own header rows are placed: a header row that moves
	// to the next page is drawn there once, not also as its own repeat (#324).
	let headersPlaced = headerRows.length === 0;
	function startNewFragment() {
		flushFragment();
		cursor.newColumn();
		if (headersPlaced) addHeaders();
	}
	function addHeaders() {
		for (const header of headerRows)
			fragmentRows.push(toRowBox(layoutRow(header, cursor.columnWidthPx, measurer, note), true));
	}

	function placeRow(row: LayoutTableRow, initial: RowLayout, cantSplit: boolean, isHeader = false) {
		const canSplitOversized = oversizedCantSplitRow !== 'clip' && !isHeader;
		let remainingRow: RowLayout | null = initial;
		while (remainingRow) {
			fragmentHeight += fragmentBorders.remove(fragmentRows.at(-1));
			const keeps =
				paragraphKeeps &&
				cursor.page.columns.length === 1 &&
				cursor.column.startYPx === undefined &&
				Math.abs(cursor.remainingHeightPx() - (cursor.columnHeightPx - cursor.usedColumnHeightPx)) <
					1e-6
					? paragraphKeeps
					: undefined;
			if (hasParagraphKeeps && !keeps && !keepWarning) {
				note(
					'Table-cell keep-lines pagination remains approximate outside the measured ordinary automatic two-cell, top-aligned, single-column, note-free row with direct exact-spaced paragraphs, zero paragraph spacing and widow control off.',
				);
				keepWarning = true;
			}
			if (
				keeps &&
				remainingRow === initial &&
				!cursor.atColumnTop &&
				keeps.startupMinimum(remainingRow) > cursor.remainingHeightPx() - fragmentHeight + 1e-6
			) {
				startNewFragment();
				continue;
			}
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
						row,
					);
					if (fragmentRows.length) {
						fragmentRows = [];
						fragmentHeight = 0;
						addHeaders();
					}
					continue;
				}
			}
			if (footnotes && !keeps) {
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
									row,
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
							row,
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
				if (remainingRow.heightPx > freshRemaining) {
					if (canSplitOversized) {
						cantSplit = false;
						continue;
					}
					note('A table row marked "keep row together" is taller than an empty page/column.');
				}
				fragmentRows.push(toRowBox(remainingRow, false));
				remainingRow = null;
				continue;
			}
			const paragraphCuts = keeps?.cuts(
				remainingRow,
				remaining,
				cursor.atColumnTop && fragmentRows.length === 0,
			);
			// On an empty page a paragraph taller than the page splits regardless of widow control.
			const widowCut =
				paragraphCuts || (cursor.atColumnTop && fragmentRows.every((item) => item.repeated))
					? undefined
					: widowControlCut(row, remainingRow, Math.max(0, remaining));
			const { before, after } = splitRowAtHeight(
				remainingRow,
				Math.max(0, Math.min(remaining, widowCut ?? Infinity)),
				undefined,
				paragraphCuts,
			);
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
			const placedPart = before.cells.some((cell) => cell.length);
			if (placedPart) fragmentRows.push(toRowBox(before, false));
			startNewFragment();
			// A row that moves whole to the next page is that row again, not a continuation: it keeps
			// its minimum height and the space before its first paragraphs (#407).
			if (!placedPart && remainingRow === initial) {
				remainingRow = layoutRow(row, cursor.columnWidthPx, measurer, note);
				continue;
			}
			// A row that moves whole to the next page keeps its merge span (a split tail has none).
			remainingRow =
				!placedPart && !paragraphCuts && previousWidth === cursor.columnWidthPx
					? remainingRow
					: after &&
						  (paragraphCuts || previousWidth !== cursor.columnWidthPx || nativeFragments.get(row))
						? layoutRow(
								{ ...row, heightPx: 0, heightRule: 'atLeast' },
								cursor.columnWidthPx,
								measurer,
								note,
								after.cells,
								row,
							)
						: after;
		}
	}

	const merged = new MergedRowPlan(table.rows);
	for (const [index, source] of table.rows.entries()) {
		merged.plan(
			index,
			(candidate) =>
				layoutRow(candidate, cursor.columnWidthPx, measurer, () => {}, undefined, source),
			cursor.remainingHeightPx() - fragmentHeight,
		);
		const row = merged.effective(source);
		if (row !== source) nativeFragments.set(row, nativeFragments.get(source) ?? false);
		const layout = layoutRow(row, cursor.columnWidthPx, measurer, note);
		if (
			layout.heightPx > cursor.remainingHeightPx() &&
			!nativeFragments.get(row) &&
			row.cells.some((cell) =>
				cell.blocks?.some((block) => block.kind === 'table' && block.rows.length > 1),
			)
		)
			note(
				'Nested physical border edges at page cuts remain approximate outside automatic, top-aligned tables with two or more two-cell, note-free rows, direct paragraphs and complete shared borders.',
			);

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
		const isHeader = headerRows.includes(source);
		placeRow(
			row,
			layout,
			Boolean(row.cantSplit) || isHeader || (!!footnotes && row.heightRule === 'exact'),
			isHeader,
		);
		if (index === headerRows.length - 1) headersPlaced = true;
	}
	flushFragment();
}

/**
 * A vertically merged cell aligns its content across every row it covers, not in its first row.
 * The first row's geometry gets the merged height when the whole merge sits in this fragment (a
 * merge broken by a page or column boundary keeps aligning within its first row).
 */
function alignMergedCells(rows: LayoutTableRowBox[]): void {
	rows.forEach((row, index) => {
		for (const geometry of row.geometry ?? []) {
			const count = geometry.mergeRows ?? 1;
			if (count < 2) continue;
			const covered = rows.slice(index, index + count);
			if (covered.length < count || covered.some((item) => item.repeated !== row.repeated))
				continue;
			shadeMergedContinuations(geometry, covered);
			if (!geometry.verticalAlign || geometry.verticalAlign === 'top') continue;
			const last = covered[count - 1]!;
			geometry.alignmentHeightPx = last.yPx + last.heightPx - row.yPx;
		}
	});
}

/** The shading of a merge's restart cell covers the whole merged cell, not its first row only. */
function shadeMergedContinuations(
	start: NonNullable<LayoutTableRowBox['geometry']>[number],
	covered: LayoutTableRowBox[],
): void {
	if (!start.shading) return;
	for (const row of covered.slice(1)) {
		const part = row.geometry?.find(
			(item) =>
				Math.abs(item.xPx - start.xPx) < 0.5 && Math.abs(item.widthPx - start.widthPx) < 0.5,
		);
		if (part && !part.shading) part.shading = start.shading;
	}
}
