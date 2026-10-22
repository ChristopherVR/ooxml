import * as schemas from '../schemas/index';
import * as chartFormattingTools from '../tools/chart-formatting-tools';
import * as chartTools from '../tools/chart-tools';
import * as chartUserShapeTools from '../tools/chart-user-shape-tools';
import * as smartartTools from '../tools/smartart-tools';
import * as templateTools from '../tools/template-tools';
import * as themeTools from '../tools/theme-tools';
import type { AutomationTool } from '../../../automation/tools/types';
import { deckTool } from './deck-tool';

/** PowerPoint tools: themes, charts, SmartArt and templates. */
export const chartsTools: readonly AutomationTool[] = [
	deckTool(
		'read',
		'get_theme_info',
		'Get current theme information (colors, fonts, available presets)',
		schemas.GetThemeInfoSchema,
		(ctx) => themeTools.getThemeInfo(ctx),
	),
	deckTool(
		'edit',
		'apply_theme_preset',
		'Apply a built-in theme preset to the presentation',
		schemas.ApplyThemePresetSchema,
		(ctx, params) => themeTools.applyThemePreset(ctx, params),
	),
	deckTool(
		'edit',
		'update_theme_colors',
		'Update individual theme colors (accent1-6, dk1/2, lt1/2, hlink, folHlink)',
		schemas.UpdateThemeColorsSchema,
		(ctx, params) => themeTools.updateThemeColors(ctx, params),
	),
	deckTool(
		'edit',
		'update_theme_fonts',
		'Update theme heading (major) and body (minor) fonts',
		schemas.UpdateThemeFontsSchema,
		(ctx, params) => themeTools.updateThemeFonts(ctx, params),
	),
	deckTool(
		'edit',
		'update_chart',
		'Update chart type, title, legend, data labels, axis, or categories',
		schemas.UpdateChartSchema,
		(ctx, params) => chartTools.updateChart(ctx, params),
	),
	deckTool(
		'edit',
		'add_chart_series',
		'Add a data series to an existing chart',
		schemas.AddChartSeriesSchema,
		(ctx, params) => chartTools.addChartSeriesT(ctx, params),
	),
	deckTool(
		'edit',
		'remove_chart_series',
		'Remove a data series from a chart by index',
		schemas.RemoveChartSeriesSchema,
		(ctx, params) => chartTools.removeChartSeriesT(ctx, params),
	),
	deckTool(
		'edit',
		'update_chart_series_data',
		'Update data values for a chart series',
		schemas.UpdateChartSeriesDataSchema,
		(ctx, params) => chartTools.updateChartSeriesData(ctx, params),
	),
	deckTool(
		'edit',
		'create_chart',
		'Create a new chart element on a slide',
		schemas.CreateChartSchema,
		(ctx, params) => chartTools.createChart(ctx, params),
	),
	deckTool(
		'edit',
		'format_chart_data_point',
		'Set per-data-point chart formatting: fill/stroke colour, stroke width, dash style (c:dPt/c:spPr), explosion, and marker override',
		schemas.FormatChartDataPointSchema,
		(ctx, params) => chartFormattingTools.formatChartDataPoint(ctx, params),
	),
	deckTool(
		'edit',
		'format_chart_data_label',
		"Set a per-data-point data-label override (c:dLbl): show flags, position, custom text, and the label's own fill/line formatting (c:spPr) and font (c:txPr)",
		schemas.FormatChartDataLabelSchema,
		(ctx, params) => chartFormattingTools.formatChartDataLabel(ctx, params),
	),
	deckTool(
		'edit',
		'format_chart_series',
		'Set a chart series marker (c:marker), trendline (c:trendline, incl. line width/dash), or error bars (c:errBars, incl. line width/dash)',
		schemas.FormatChartSeriesSchema,
		(ctx, params) => chartFormattingTools.formatChartSeries(ctx, params),
	),
	deckTool(
		'edit',
		'set_chart_helper_line',
		'Set or remove a chart-level helper line on a line/stock/combo chart: c:dropLines or c:hiLowLines',
		schemas.SetChartHelperLineSchema,
		(ctx, params) => chartFormattingTools.setChartHelperLineT(ctx, params),
	),
	deckTool(
		'edit',
		'set_chart_color_map_override',
		"Set or remove a chart's theme colour-map override (c:clrMapOvr), remapping theme colour roles for this chart only",
		schemas.SetChartColorMapOverrideSchema,
		(ctx, params) => chartFormattingTools.setChartColorMapOverrideT(ctx, params),
	),
	deckTool(
		'read',
		'chart_user_shape_list',
		"List a chart's drawing-overlay shapes (c:userShapes)",
		schemas.ListChartUserShapesSchema,
		(ctx, params) => chartUserShapeTools.listChartUserShapesT(ctx, params),
	),
	deckTool(
		'edit',
		'chart_user_shape_add',
		"Add a drawing-overlay shape (text box or connector) to a chart's c:userShapes",
		schemas.AddChartUserShapeSchema,
		(ctx, params) => chartUserShapeTools.addChartUserShapeT(ctx, params),
	),
	deckTool(
		'edit',
		'chart_user_shape_update',
		"Move, resize, or restyle one of a chart's drawing-overlay shapes by index",
		schemas.UpdateChartUserShapeSchema,
		(ctx, params) => chartUserShapeTools.updateChartUserShapeT(ctx, params),
	),
	deckTool(
		'edit',
		'chart_user_shape_remove',
		"Remove one of a chart's drawing-overlay shapes by index",
		schemas.RemoveChartUserShapeSchema,
		(ctx, params) => chartUserShapeTools.removeChartUserShapeT(ctx, params),
	),
	deckTool(
		'edit',
		'manage_smart_art',
		'Manage SmartArt: get nodes, add/remove/reorder/promote/demote nodes, or decompose to shapes',
		schemas.ManageSmartArtSchema,
		(ctx, params) => smartartTools.manageSmartArt(ctx, params),
		{ returnsText: true },
	),
	deckTool(
		'read',
		'find_placeholders',
		'Discover all {{placeholder}} tokens in the presentation',
		schemas.FindPlaceholdersSchema,
		(ctx) => templateTools.findPlaceholdersT(ctx),
	),
	deckTool(
		'edit',
		'apply_template',
		'Replace {{placeholder}} tokens with provided values (mail merge)',
		schemas.ApplyTemplateSchema,
		(ctx, params) =>
			templateTools.applyTemplateT(ctx, { data: params.data as Record<string, unknown> }),
	),
];
