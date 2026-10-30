// Registers the tool dialogs: charts, paste special, page setup, zoom, sheets, protection, cell
// shifting, data tools, symbols, series and tables.
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { openCellShift, openOutlineAxis } from './cell-shift';
import { type CreateTableProps, openCreateTable } from './create-table';
import { type CustomAutoFilterProps, openCustomAutoFilter } from './custom-autofilter';
import { type Top10Props, openTop10Filter } from './top10-filter';
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
import { openDataTable, openGoalSeek } from './what-if';
import { openForecastSheet } from './forecast-sheet';
import { openScenarioManager } from './scenario-manager';
import { openCustomViews } from './custom-views';
import { openCreateSparklines, openSparklineSettings } from './sparklines';
import { openPivotFieldSettings } from './pivot-field-settings';
import { openPivotGroup } from './pivot-group';
import { openPivotGroupDates } from './pivot-group-dates';
import { openPivotTable } from './pivots';
import {
	openInsertSlicer,
	openInsertTimeline,
	openSlicerSettings,
	openTimelineSettings,
} from './slicers';
import { openPrintPreview } from './print-preview';
import type { SparklineType } from 'ooxml-core/xlsx';

const props = <T>(value: unknown): T =>
	value && typeof value === 'object' ? (value as T) : ({} as T);

export function registerToolDialogs(ctx: EditorContext): void {
	const d = ctx.dialogs;
	d.register('print-preview', (c) => openPrintPreview(c));
	d.register('create-pivot-table', (c) => openPivotTable(c));
	d.register('pivot-settings', (c) => openPivotTable(c, true));
	d.register('pivot-group', (c) => openPivotGroup(c));
	d.register('pivot-group-dates', (c) => openPivotGroupDates(c));
	d.register('pivot-field-settings', (c) => openPivotFieldSettings(c));
	d.register('insert-slicer', (c) => openInsertSlicer(c));
	d.register('slicer-settings', (c) => openSlicerSettings(c));
	d.register('insert-timeline', (c) => openInsertTimeline(c));
	d.register('timeline-settings', (c) => openTimelineSettings(c));
	d.register('create-sparklines', (c, p) =>
		openCreateSparklines(c, props<{ type?: SparklineType }>(p).type),
	);
	d.register('sparkline-settings', (c) => openSparklineSettings(c));
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
	d.register('custom-autofilter', (c, p) =>
		openCustomAutoFilter(c, props<CustomAutoFilterProps>(p)),
	);
	d.register('top10-filter', (c, p) => openTop10Filter(c, props<Top10Props>(p)));
	d.register('remove-duplicates', (c) => openRemoveDuplicates(c));
	d.register('text-to-columns', (c) => openTextToColumns(c));
	d.register('goal-seek', (c) => openGoalSeek(c));
	d.register('data-table', (c) => openDataTable(c));
	d.register('forecast-sheet', (c) => openForecastSheet(c));
	d.register('scenario-manager', (c) => openScenarioManager(c));
	d.register('custom-views', (c) => openCustomViews(c));
	d.register('symbol', (c) => openSymbol(c));
	d.register('fill-series', (c) => openFillSeries(c));
	d.register('create-table', (c, p) => openCreateTable(c, props<CreateTableProps>(p)));
}
