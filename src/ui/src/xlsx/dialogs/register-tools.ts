// Registers the tool dialogs: charts, paste special, page setup, zoom, sheets, protection, cell
// shifting, data tools, symbols, series and tables.
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { openCellShift, openOutlineAxis } from './cell-shift';
import { type CreateTableProps, openCreateTable } from './create-table';
import { openFillSeries } from './fill-series';
import { type InsertChartProps, openInsertChart } from './insert-chart';
import { openMoveCopySheet } from './move-copy-sheet';
import { type PageSetupProps, openPageSetup } from './page-setup';
import { openPasteSpecial } from './paste-special';
import { openProtectSheet } from './protect-sheet';
import { openRemoveDuplicates } from './remove-duplicates';
import { openSymbol } from './symbol';
import { openTextToColumns } from './text-to-columns';
import { openZoom } from './zoom';

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
