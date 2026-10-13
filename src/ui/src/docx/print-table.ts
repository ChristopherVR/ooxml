import { cellVerticalOffset } from 'ooxml-core/docx/layout';
import type {
	LayoutBorder,
	LayoutCellGeometry,
	LayoutParagraphBox,
	LayoutTableBox,
	LayoutTableRowBox,
} from 'ooxml-core/docx/layout';

const cssBorder = (border: LayoutBorder | undefined) =>
	border ? `${border.widthPx}px ${border.style} ${border.color}` : '';

/** Rows without geometry (engine-level input) share the table width evenly. */
function geometryOf(row: LayoutTableRowBox, index: number, widthPx: number): LayoutCellGeometry {
	const given = row.geometry?.[index];
	if (given) return given;
	const share = widthPx / Math.max(1, row.cells.length);
	return {
		xPx: share * index,
		widthPx: share,
		paddingLeftPx: 0,
		paddingRightPx: 0,
		contentHeightPx: 0,
	};
}

/**
 * Draws one table fragment: cells at their grid positions with shading and the document's
 * borders (no gridlines, as Word prints), content inset by cell margins and aligned vertically.
 * Each cell draws its top and left edges; the last cell of a row adds its right edge and the
 * fragment's last row its bottom, so shared edges are drawn once.
 */
export function renderTable(
	box: LayoutTableBox,
	columnWidthPx: number,
	renderParagraph: (paragraph: LayoutParagraphBox) => HTMLElement | null,
): HTMLElement {
	const table = document.createElement('div');
	table.className = 'dve-print-block dve-print-table';
	table.style.top = `${box.yPx}px`;
	table.style.height = `${box.heightPx}px`;
	if (box.widthPx !== undefined) table.style.width = `${box.widthPx}px`;
	if (box.xPx) table.style.left = `${box.xPx}px`;
	box.rows.forEach((row, rowIndex) => {
		const rowEl = document.createElement('div');
		rowEl.className = 'dve-print-row';
		rowEl.style.top = `${row.yPx}px`;
		rowEl.style.height = `${row.heightPx}px`;
		if (row.repeated) rowEl.dataset.repeated = 'true';
		const lastRow = rowIndex === box.rows.length - 1;
		row.cells.forEach((paragraphs, cellIndex) => {
			const geometry = geometryOf(row, cellIndex, columnWidthPx);
			const cell = document.createElement('div');
			cell.className = 'dve-print-cell';
			Object.assign(cell.style, {
				left: `${geometry.xPx}px`,
				width: `${geometry.widthPx}px`,
				height: `${row.heightPx}px`,
			});
			if (geometry.shading) cell.style.background = geometry.shading;
			const borders = geometry.borders;
			cell.style.borderTop = cssBorder(borders?.top);
			cell.style.borderLeft = cssBorder(borders?.left);
			if (cellIndex === row.cells.length - 1) cell.style.borderRight = cssBorder(borders?.right);
			if (lastRow) cell.style.borderBottom = cssBorder(borders?.bottom);
			const content = document.createElement('div');
			content.className = 'dve-print-cell-content';
			const offset = cellVerticalOffset(geometry, row.heightPx);
			Object.assign(content.style, {
				left: `${geometry.paddingLeftPx - (geometry.borderInsetsReserved ? (borders?.left?.widthPx ?? 0) : 0)}px`,
				right: `${geometry.paddingRightPx - (geometry.borderInsetsReserved && cellIndex === row.cells.length - 1 ? (borders?.right?.widthPx ?? 0) : 0)}px`,
				top: `${offset - (geometry.borderInsetsReserved ? (borders?.top?.widthPx ?? 0) : 0)}px`,
			});
			for (const nested of row.cellTables?.[cellIndex] ?? []) {
				const frame = renderTable(nested, nested.widthPx ?? geometry.widthPx, () => null);
				frame.classList.add('dve-print-nested-table');
				frame.setAttribute('aria-hidden', 'true');
				frame.style.pointerEvents = 'none';
				content.append(frame);
			}
			for (const paragraph of paragraphs) {
				const element = renderParagraph(paragraph);
				if (element) content.append(element);
			}
			cell.append(content);
			rowEl.append(cell);
		});
		table.append(rowEl);
	});
	return table;
}
