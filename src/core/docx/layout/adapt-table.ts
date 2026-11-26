import { rowSpanOf, tableGrid } from '../table-grid';
import {
	resolveCellVisuals,
	resolveTableStyleFormatting,
	resolveThemeColorReference,
	twipsToPixels,
	type DocumentModel,
	type Paragraph,
	type RunFormatting,
	type Table,
	type TableBorderSide,
} from '../index';
import { expectDefined } from '../expect-defined';
import { autofitOverhangsCellMargins, compatibilityProfile } from '../compatibility';
import { floatingFrame } from './adapt-table-floating';
import { resolveTableCellMargins } from '../table-margins';
import { cellSourceBlocks } from '../cell-blocks';
import type {
	LayoutBorder,
	LayoutCellBorders,
	LayoutCellPadding,
	LayoutParagraph,
	LayoutTable,
	LayoutTableCell,
	LayoutTableRow,
} from './input';

const twipsToPx = twipsToPixels;

/** `#rrggbb` from a model color, which the parser stores with or without its `#`. */
export function cssHex(value: string | undefined): string | undefined {
	const hex = value?.replace(/^#/, '');
	return hex && /^[0-9a-f]{6}$/i.test(hex) ? `#${hex}` : undefined;
}

export function border(
	side: TableBorderSide | undefined,
	theme: DocumentModel['theme'],
	preserveSingleWidth = false,
): LayoutBorder | undefined {
	if (!side?.style || side.style === 'none' || side.style === 'nil') return undefined;
	const style: LayoutBorder['style'] = side.style.startsWith('double')
		? 'double'
		: /dot/i.test(side.style)
			? 'dotted'
			: /dash/i.test(side.style)
				? 'dashed'
				: 'solid';
	const themed =
		side.themeColor && theme
			? resolveThemeColorReference({ token: side.themeColor }, theme)
			: undefined;
	const hex = cssHex(side.color);
	// `sz` is in eighths of a point; a double line needs room for both strokes.
	const widthPx = Math.max(
		style === 'double' ? 3 : preserveSingleWidth && style === 'solid' ? 0 : 1,
		(((side.sizeEighthPoints ?? 4) / 8) * 96) / 72,
	);
	return { widthPx, style, color: themed ?? hex ?? '#000000' };
}

/** Row height, keep-together and repeat-as-header flags for pagination. */
function rowLayout(
	props: NonNullable<Table['rowProperties']>[number] | undefined,
): Partial<LayoutTableRow> {
	if (!props) return {};
	return {
		...(props.heightTwips !== undefined ? { heightSpecified: true } : {}),
		...(props.heightTwips
			? { heightPx: twipsToPx(props.heightTwips), heightRule: props.heightRule ?? 'atLeast' }
			: {}),
		...(props.cantSplit ? { cantSplit: true } : {}),
		...(props.header ? { isHeader: true } : {}),
	};
}

/**
 * Converts a model table for pagination: grid positions and widths (merged cells span grid
 * columns), cell margins, resolved borders and shading (table style, table and cell layers), and
 * vertical merges drawn as one cell by hiding the borders between the merged cells.
 */
export function adaptTable(
	table: Table,
	model: DocumentModel,
	adaptParagraph: (
		paragraph: Paragraph,
		markLabel?: string,
		tableStyleRun?: RunFormatting,
	) => LayoutParagraph,
	_nested = false,
): LayoutTable {
	const grid = table.grid?.map(twipsToPx);
	const columnCount =
		grid?.length ??
		Math.max(
			1,
			...table.rows.map((row) => row.reduce((width, cell) => width + (cell.gridSpan ?? 1), 0)),
		);
	const gridWidth = (start: number, span: number) =>
		grid ? grid.slice(start, start + span).reduce((sum, width) => sum + width, 0) : undefined;
	// Grid column of each cell, so vertical merges line up across rows with different spans.
	const gridIndexes = table.rows.map((row) => {
		let index = 0;
		return row.map((cell) => {
			const start = index;
			index += cell.gridSpan ?? 1;
			return start;
		});
	});
	const merges = tableGrid(table);
	// Default cell margins: the table's own, its style chain, then the default table style.
	const tableMargins = resolveTableCellMargins(table, model.tableStyles);
	const cellAt = (rowIndex: number, gridIndex: number) => {
		const position = gridIndexes[rowIndex]?.indexOf(gridIndex) ?? -1;
		return position >= 0 ? table.rows[rowIndex]?.[position] : undefined;
	};
	return {
		kind: 'table',
		id: table.id,
		...(grid ? { widthPx: grid.reduce((sum, width) => sum + width, 0) } : {}),
		...(table.indentTwips ? { indentPx: twipsToPx(table.indentTwips) } : {}),
		...(table.alignment ? { alignment: table.alignment } : {}),
		...(table.layout === 'autofit'
			? {
					autofit: {
						columns: Math.max(
							columnCount,
							...table.rows.map((row) => row.reduce((sum, cell) => sum + (cell.gridSpan ?? 1), 0)),
						),
						...(table.widthTwips ? { widthPx: twipsToPx(table.widthTwips) } : {}),
						...(table.widthPercent ? { widthFraction: table.widthPercent } : {}),
						...(autofitOverhangsCellMargins(compatibilityProfile(model))
							? { legacyMargins: true }
							: {}),
					},
				}
			: {}),
		...(table.floating && !_nested ? { floating: floatingFrame(table.floating) } : {}),
		rows: table.rows.map((row, rowIndex) => ({
			...rowLayout(table.rowProperties?.[rowIndex]),
			cells: row.map((cell, cellIndex): LayoutTableCell => {
				const start = expectDefined(gridIndexes[rowIndex]?.[cellIndex], 'table cell grid index');
				const span = cell.gridSpan ?? 1;
				const visuals = resolveCellVisuals(
					table,
					cell,
					{
						row: rowIndex,
						lastRow: rowIndex,
						column: start,
						lastColumn: start + span - 1,
						rowCount: table.rows.length,
						columnCount,
					},
					model.tableStyles,
					model.theme,
				);
				const borders: LayoutCellBorders = {};
				for (const side of ['top', 'right', 'bottom', 'left'] as const) {
					const resolved = border(visuals.borders?.[side], model.theme, true);
					if (resolved) borders[side] = resolved;
				}
				if (cell.verticalMerge === 'continue') delete borders.top;
				if (cellAt(rowIndex + 1, start)?.verticalMerge === 'continue') delete borders.bottom;
				const margins = { ...tableMargins, ...cell.margins };
				const padding: LayoutCellPadding = {
					top: twipsToPx(margins.top),
					right: twipsToPx(margins.right),
					bottom: twipsToPx(margins.bottom),
					left: twipsToPx(margins.left),
				};
				const width =
					gridWidth(start, span) ?? (cell.widthTwips ? twipsToPx(cell.widthTwips) : undefined);
				const xPx = grid ? gridWidth(0, start) : undefined;
				const fill = cssHex(visuals.shadingFill);
				const entry = merges?.rows[rowIndex]?.[cellIndex];
				const rowSpan = merges && entry ? rowSpanOf(merges, entry) : 1;
				const blocks = cellSourceBlocks(cell);
				const last = blocks.at(-1);
				const tableStyleRun = resolveTableStyleFormatting(
					table.style,
					model.tableStyles,
					table.look,
					rowIndex,
					table.rows.length,
					start,
					columnCount,
				).run;
				const adaptCellParagraph = (paragraph: Paragraph): LayoutParagraph => {
					const adapted = adaptParagraph(paragraph, undefined, tableStyleRun);
					return cell.hideMark && last?.id === paragraph.id
						? { ...adapted, hideCellEndMark: true }
						: adapted;
				};
				return {
					paragraphs: cell.paragraphs.map(adaptCellParagraph),
					...(cell.gridSpan !== undefined ? { gridSpan: cell.gridSpan } : {}),
					...(cell.verticalMerge ? { verticalMerge: cell.verticalMerge } : {}),
					...(rowSpan > 1 ? { rowSpan } : {}),
					...(cell.sourceBlocks
						? {
								blocks: blocks.map((block) =>
									block.type === 'paragraph'
										? adaptCellParagraph(block)
										: adaptTable(block, model, adaptParagraph, true),
								),
							}
						: {}),
					...(width !== undefined ? { widthPx: width } : {}),
					...(grid && xPx !== undefined ? { xPx } : {}),
					...(cell.widthTwips ? { preferredWidthPx: twipsToPx(cell.widthTwips) } : {}),
					...(cell.widthPercent ? { preferredFraction: cell.widthPercent } : {}),
					padding,
					...(Object.keys(borders).length ? { borders } : {}),
					...(fill ? { shading: fill } : {}),
					// `both` (vertically justified) has no layout support yet and renders as top, as before.
					...(cell.verticalAlign
						? { verticalAlign: cell.verticalAlign === 'both' ? 'top' : cell.verticalAlign }
						: {}),
				};
			}),
		})),
	};
}
