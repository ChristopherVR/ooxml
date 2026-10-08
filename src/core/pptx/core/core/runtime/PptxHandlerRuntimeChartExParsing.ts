/**
 * @fileoverview ChartEx (`cx:chartSpace`) parsing and the chart-part facts both chart families
 * read the same way: the external workbook and its embedded rows, pivot source and formats, the
 * Office 2013+ colour-style and chart-style parts, layouts, user shapes, print settings and
 * protection. ChartEx parts are not modelled by the neutral `chart` area, so they stay on the
 * `fast-xml-parser` object tree (`parseCxChartSeries`).
 *
 * Mixin chain position:
 *   `legacy-chart-parsing` -> **this** -> `PptxHandlerRuntimeChartParsing`
 */

import { XmlObject } from '../../types';
import type { PptxChartData } from '../../types';
import { parseCxChartSeries } from '../../utils/chart-cx-parser';
import { parseChartLayouts } from '../../utils/chart-layout';
import { parseChartPivotFormats } from '../../utils/chart-pivot-formats';
import { parseChartPrintSettings } from '../../utils/chart-print-settings';
import { parseChartProtection } from '../../utils/chart-protection';
import { parseChartSpaceFlags } from '../../utils/chart-space-flags';
import { PptxHandlerRuntime as PptxHandlerRuntimeBase } from './legacy-chart-parsing';

/** The chart-part facts {@link PptxHandlerRuntime.readChartPartFacts} returns. */
type ChartPartFacts = Pick<
	PptxChartData,
	| 'externalData'
	| 'embeddedWorkbookData'
	| 'pivotSource'
	| 'colorPalette'
	| 'colorMethod'
	| 'colorStylePartPath'
	| 'colorStyleOriginalPalette'
	| 'colorStyleOriginalMethod'
	| 'chartStyleDefinition'
	| 'layouts'
	| 'userShapesXml'
	| 'userShapes'
	| 'pivotFormats'
	| 'printSettings'
	| 'protection'
>;

export class PptxHandlerRuntime extends PptxHandlerRuntimeBase {
	/**
	 * The facts of a chart part that live outside the plot (or in other parts): the external data
	 * reference and the embedded workbook's rows, pivot source and formats, the colour-style and
	 * chart-style parts, layouts, user shapes, print settings and protection. Only the facts that
	 * are present are set.
	 */
	protected async readChartPartFacts(
		chartSpace: XmlObject | undefined,
		chartRoot: XmlObject | undefined,
		chartPartPath: string,
	): Promise<ChartPartFacts> {
		const localName = (key: string) => this.compatibilityService.getXmlLocalName(key);
		const externalData = await this.parseChartExternalData(chartSpace, chartPartPath);
		const embeddedWorkbookData = await this.parseEmbeddedWorkbook(externalData);
		const pivotSource = this.parsePivotSource(chartSpace);
		// Office 2013+ colour style (colorsN.xml) and the separate chart style part (styleN.xml),
		// which carries the built-in style's per-element font, line and fill defaults.
		const chartColorStyle = await this.parseChartColorStyle(chartPartPath);
		const chartStyleDefinition = await this.parseChartStyleDefinitionPart(chartPartPath);
		const layouts = parseChartLayouts(chartRoot, localName);
		const userShapesXml = this.parseUserShapesXml(chartSpace);
		const userShapes = await this.parseChartUserShapes(chartSpace, chartPartPath);
		const pivotFormats = parseChartPivotFormats(chartRoot, localName, {
			parseColor: (node, placeholder) => this.parseColor(node, placeholder),
		});
		const printSettings = parseChartPrintSettings(chartSpace, localName);
		const protection = parseChartProtection(chartSpace, localName);
		return {
			...(externalData ? { externalData } : {}),
			...(embeddedWorkbookData ? { embeddedWorkbookData } : {}),
			...(pivotSource ? { pivotSource } : {}),
			...(chartColorStyle?.palette ? { colorPalette: chartColorStyle.palette } : {}),
			...(chartColorStyle?.method ? { colorMethod: chartColorStyle.method } : {}),
			...(chartColorStyle
				? {
						colorStylePartPath: chartColorStyle.partPath,
						colorStyleOriginalPalette: [...chartColorStyle.palette],
						colorStyleOriginalMethod: chartColorStyle.method,
					}
				: {}),
			...(layouts ? { layouts } : {}),
			...(userShapesXml ? { userShapesXml } : {}),
			...(userShapes ? { userShapes } : {}),
			...(pivotFormats ? { pivotFormats } : {}),
			...(printSettings ? { printSettings } : {}),
			...(protection ? { protection } : {}),
			...(chartStyleDefinition ? { chartStyleDefinition } : {}),
		};
	}

	/**
	 * Resolve a chart title's text when it was authored as a linked cell reference
	 * (`c:title/c:tx/c:strRef`, PowerPoint's "Title Linked to Cell") rather than rich text: its
	 * cached text lives in `c:strRef/c:strCache/c:pt/c:v`.
	 */
	protected resolveChartLinkedTitleText(titleNode: XmlObject | undefined): string | undefined {
		const titleTx = this.xmlLookupService.getChildByLocalName(titleNode, 'tx');
		const cached = this.extractChartPointValues(titleTx, false);
		return cached[0]?.trim() || undefined;
	}

	/**
	 * Parse a cx: namespace (Office 2016+) chart using the utility parser.
	 *
	 * @param plotArea - The `c:plotArea` XML object.
	 * @param chartType - The detected chart type.
	 * @param chartSpace - The `c:chartSpace` XML root.
	 * @param chartRoot - The `c:chart` XML element.
	 * @param chartPartPath - The ZIP path of the chart part.
	 * @param chartRelationshipId - The relationship ID linking the slide to this chart.
	 * @returns Parsed chart data, or `undefined` if cx parsing yields no series.
	 */
	protected async parseCxChart(
		plotArea: XmlObject,
		chartType: PptxChartData['chartType'],
		chartSpace: XmlObject | undefined,
		chartRoot: XmlObject | undefined,
		chartPartPath: string,
		chartRelationshipId: string,
	): Promise<PptxChartData | undefined> {
		const result = parseCxChartSeries(
			plotArea,
			this.xmlLookupService,
			chartSpace,
			chartRoot,
			{ parseColor: (node, placeholder) => this.parseColor(node, placeholder) },
			(raw) => this.resolveThemeTypeface(raw) ?? raw,
		);
		if (!result) {
			return undefined;
		}
		// The ChartEx (`cx:`) series parser may attach a `chartData` override (axes, title, and
		// any waterfall/histogram-specific fields resolved by its own sub-parsers) that should win
		// over this function's own title/style/axes resolution.
		const cxOverride = result.chartData;

		const titleNode = this.xmlLookupService.getChildByLocalName(chartRoot, 'title');
		const titleTextValues: string[] = [];
		this.collectLocalTextValues(titleNode, 't', titleTextValues);
		const cxTitleText =
			cxOverride?.title ?? titleTextValues[0] ?? this.resolveChartLinkedTitleText(titleNode);
		const chartStyle = this.extractChartStyle(chartSpace, chartRoot);

		// Merge hasDataLabels from cx: data labels parsing
		if (result.hasDataLabels && chartStyle) {
			chartStyle.hasDataLabels = true;
		}

		// Parse plotVisOnly (c:plotVisOnly): defaults to true when absent
		const plotVisibleOnly = this.parsePlotVisOnly(chartRoot);
		const facts = await this.readChartPartFacts(chartSpace, chartRoot, chartPartPath);
		const chartSpaceFlags = parseChartSpaceFlags(chartSpace, this.xmlLookupService);
		const view3D = this.parseView3D(chartRoot);
		const chartChrome = this.parseChartChrome(chartRoot);
		const clrMapOvr = this.parseClrMapOvr(chartSpace);

		return {
			chartType,
			categories: result.categories,
			...(result.categoryLevels ? { categoryLevels: result.categoryLevels } : {}),
			series: result.series,
			title: cxTitleText,
			style: chartStyle,
			chartPartPath,
			chartRelationshipId,
			...facts,
			...(plotVisibleOnly !== undefined ? { plotVisibleOnly } : {}),
			...(view3D ? { view3D } : {}),
			...(chartChrome ? { chartChrome } : {}),
			...(clrMapOvr ? { clrMapOvr } : {}),
			...(chartSpaceFlags.date1904 !== undefined ? { date1904: chartSpaceFlags.date1904 } : {}),
			...(chartSpaceFlags.roundedCorners !== undefined
				? { roundedCorners: chartSpaceFlags.roundedCorners }
				: {}),
			// Placed last so a ChartEx sub-parser's own override (title, style,
			// axes, ...) wins over this function's generic resolution above.
			...(cxOverride ?? {}),
		};
	}
}
