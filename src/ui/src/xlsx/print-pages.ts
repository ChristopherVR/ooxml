import {
	cellView,
	createGridMetrics,
	createPrintMergeIndex,
	formatAddress,
	paginateWorksheet,
	type PrintPage,
	type Workbook,
} from 'ooxml-core/xlsx';
import { cellCss, escapeHtml } from './print-cell';
import { printPageHeaders, type PrintHeaderOptions } from './print-header';

function pageTable(workbook: Workbook, sheetIndex: number, page: PrintPage): string {
	const sheet = workbook.sheets[sheetIndex]!;
	const metrics = createGridMetrics(sheet, { zoom: 100 });
	const headings = !!sheet.printOptions?.headings;
	const merges = createPrintMergeIndex(sheet.merges, page.rows, page.columns);
	const columns = `${headings ? '<col style="width:24px">' : ''}${page.columns.map((col) => `<col style="width:${metrics.colWidth(col)}px">`).join('')}`;
	const header = headings
		? `<tr style="height:24px"><th></th>${page.columns.map((col) => `<th>${formatAddress({ row: 0, col }).replace(/\d+$/, '')}</th>`).join('')}</tr>`
		: '';
	const rows = page.rows
		.map((row) => {
			const cells = page.columns
				.map((col) => {
					const merge = merges.at(row, col);
					if (merge && (row !== merge.row || col !== merge.col)) return '';
					const view = cellView(
						workbook,
						sheetIndex,
						merge?.range.start.row ?? row,
						merge?.range.start.col ?? col,
						undefined,
						{ merged: !!merge },
					);
					let css = cellCss(view);
					if (sheet.printOptions?.gridLines)
						for (const side of ['top', 'right', 'bottom', 'left'] as const)
							if (!view.borders[side]) css += `;border-${side}:1px solid #d9d9d9`;
					const span = `${merge && merge.colspan > 1 ? ` colspan="${merge.colspan}"` : ''}${merge && merge.rowspan > 1 ? ` rowspan="${merge.rowspan}"` : ''}`;
					return `<td${span} style="${escapeHtml(css)}">${escapeHtml(view.text)}</td>`;
				})
				.join('');
			return `<tr style="height:${metrics.rowHeight(row)}px">${headings ? `<th>${row + 1}</th>` : ''}${cells}</tr>`;
		})
		.join('');
	return `<table style="width:${page.width + (headings ? 24 : 0)}px"><colgroup>${columns}</colgroup><tbody>${header}${rows}</tbody></table>`;
}

/** Browser preview and printing use the exact same core page plan. */
export function buildPrintHtml(
	workbook: Workbook,
	sheetIndex: number,
	title: string,
	options: PrintHeaderOptions = {},
): string {
	const headerOptions = { date: new Date(), ...options };
	const layout = paginateWorksheet(workbook, sheetIndex);
	const sheet = workbook.sheets[sheetIndex]!;
	const pages = layout.pages
		.map((page) => {
			const headings = sheet.printOptions?.headings ? 24 : 0;
			const left = sheet.printOptions?.horizontalCentered
				? Math.max(0, (layout.contentWidth - (page.width + headings) * layout.scale) / 2)
				: 0;
			const top = sheet.printOptions?.verticalCentered
				? Math.max(0, (layout.contentHeight - (page.height + headings) * layout.scale) / 2)
				: 0;
			return `<section class="xve-print-page" data-page="${page.number}" style="position:relative;width:${layout.paperWidth}px;height:${layout.paperHeight}px;padding:${layout.margins.top}in ${layout.margins.right}in ${layout.margins.bottom}in ${layout.margins.left}in">${printPageHeaders(workbook, sheetIndex, layout, page.number, title, headerOptions)}<div class="xve-print-content" style="width:${layout.contentWidth}px;height:${layout.contentHeight}px"><div style="position:absolute;left:${left}px;top:${top}px;transform:scale(${layout.scale});transform-origin:top left">${pageTable(workbook, sheetIndex, page)}</div></div></section>`;
		})
		.join('');
	return `<!doctype html><html><head><meta charset="utf-8"><title>${escapeHtml(title)}</title><style>
@page{size:${layout.paperWidth / 96}in ${layout.paperHeight / 96}in;margin:0}
body{margin:0;background:#e8eaed;font-family:Calibri,sans-serif}.xve-print-page{box-sizing:border-box;background:white;margin:16px auto;box-shadow:0 1px 5px #0003;overflow:hidden;break-after:page}.xve-print-content{position:relative;overflow:hidden}
table{border-collapse:collapse;table-layout:fixed}td{overflow:hidden;padding:0 3px;box-sizing:border-box;line-height:1.2}th{box-sizing:border-box;font-size:12px;font-weight:400;color:#444;border:1px solid #ccc;overflow:hidden}
@media print{body{background:white}.xve-print-page{margin:0;box-shadow:none}.xve-print-page:last-child{break-after:auto}}
</style></head><body>${pages}</body></html>`;
}
