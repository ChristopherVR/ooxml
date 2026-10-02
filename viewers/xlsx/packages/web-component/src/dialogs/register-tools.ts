// Registers the tool dialogs: charts, paste special, page setup, zoom, sheets, protection, cell
// shifting, data tools, symbols, series and tables.
import type { EditorContext } from '../context.js';
import { openCellShift, openOutlineAxis } from './cell-shift.js';
import { type CreateTableProps, openCreateTable } from './create-table.js';
import { openFillSeries } from './fill-series.js';
import { type InsertChartProps, openInsertChart } from './insert-chart.js';
import { openMoveCopySheet } from './move-copy-sheet.js';
import { type PageSetupProps, openPageSetup } from './page-setup.js';
import { openPasteSpecial } from './paste-special.js';
import { openProtectSheet } from './protect-sheet.js';
import { openRemoveDuplicates } from './remove-duplicates.js';
import { openSymbol } from './symbol.js';
import { openTextToColumns } from './text-to-columns.js';
import { openZoom } from './zoom.js';

const props = <T>(value: unknown): T =>
	value && typeof value === 'object' ? (value as T) : ({} as T);

export function registerToolDialogs(ctx: EditorContext): void {
	const d = ctx.dialogs;
	d.register('insert-chart', (c, p) => openInsertChart(c, props<InsertChartProps>(p)));
	d.register('paste-special', (c) => openPasteSpecial(c));
	d.register('page-setup', (c, p) => openPageSetup(c, props<PageSetupProps>(p)));
	d.register('zoom', (c) => openZoom(c));
	d.register('move-copy-sheet', (c) => openMoveCopySheet(c));
	d.register('protect-sheet', (c) => openProtectSheet(c));
	d.register('insert-cells', (c) => openCellShift(c, 'insert'));
	d.register('delete-cells', (c) => openCellShift(c, 'delete'));
	d.register('group', (c) => openOutlineAxis(c, 'group'));
	d.register('ungroup', (c) => openOutlineAxis(c, 'ungroup'));
	d.register('remove-duplicates', (c) => openRemoveDuplicates(c));
	d.register('text-to-columns', (c) => openTextToColumns(c));
	d.register('symbol', (c) => openSymbol(c));
	d.register('fill-series', (c) => openFillSeries(c));
	d.register('create-table', (c, p) => openCreateTable(c, props<CreateTableProps>(p)));
}
