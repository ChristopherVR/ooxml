/**
 * @fileoverview Chart parsing orchestrator for OOXML chart graphic frames.
 *
 * A classic chart part (`c:chartSpace`) is read by the neutral parser (`chart/parseChartSpace`),
 * and `PptxChartData` is built from that model through the adapter in
 * `utils/chart-from-neutral*.ts` (docs/agnostic-core-plan.md, steps 5 and 6). What the neutral
 * model does not carry (resolved colours and fonts, gradients, trendlines, error bars, picture
 * fills, c15/c16 extensions, layouts, the colour-style, chart-style, workbook and user-shape
 * parts, `c:clrMapOvr`, pivot source and formats, print settings, protection) is still read from
 * the part's `fast-xml-parser` tree by the existing object-tree readers. ChartEx parts stay on the
 * object tree entirely (`PptxHandlerRuntimeChartExParsing`).
 *
 * Mixin chain position:
 *   `PptxHandlerRuntimeChartExParsing` -> **this** -> `PptxHandlerRuntimePresentationStructure`
 */

import type { ChartSpace } from '../../../../chart/index';
import { XmlObject } from '../../types';
import type { PptxChartData } from '../../types';
import { parseLineStyle } from '../../utils/chart-advanced-parser';
import { parseChartAxes, parseChart3DSurfaces } from '../../utils/chart-axis-parser';
import { parseChartBandFmts } from '../../utils/chart-band-fmts';
import { parseBubbleChartOptions } from '../../utils/chart-bubble-options';
import { chartContainerAllows } from '../../utils/chart-container-content-model';
import { parseDataTable } from '../../utils/chart-data-table-parser';
import { parseFilteredTitles } from '../../utils/chart-ext-titles';
import { parseFilteredSeries } from '../../utils/chart-filtered-series';
import {
	axesFromNeutral,
	chartChromeFromNeutral,
	chartStyleFromNeutral,
	chartTypeFromNeutral,
	groupChartType,
	groupOptionsFromNeutral,
	neutralCategories,
	neutralDateCategories,
	neutralTitleText,
	pairGroupNodes,
	rawChildren,
	seriesFromNeutral,
	view3DFromNeutral,
	type ChartTreeReader,
} from '../../utils/chart-from-neutral';
import { readChartPartModel } from '../../utils/chart-part-repair';
import { parseChartTitleRuns } from '../../utils/chart-title-runs-parser';
import { parseChartUpDownBars } from '../../utils/chart-up-down-bars';
import { resolveDataPointPictureImages } from './chart-datapoint-picture-resolver';
import { applyEmbeddedWorkbookFallback } from './chart-embedded-fallback';
import { PptxHandlerRuntime as PptxHandlerRuntimeBase } from './PptxHandlerRuntimeChartExParsing';

/** The object-tree nodes of a chart part. */
interface ChartTree {
	chartSpace: XmlObject | undefined;
	chartRoot: XmlObject | undefined;
	plotArea: XmlObject;
}

export class PptxHandlerRuntime extends PptxHandlerRuntimeBase {
	/**
	 * Parse chart data from a graphic frame element on a slide.
	 *
	 * Resolves the chart relationship, reads the chart part, builds the chart model from the
	 * neutral `c:chartSpace` model (classic parts) or the ChartEx reader, and gathers the part's
	 * metadata (axes, data table, external data, colour style...) into one {@link PptxChartData}.
	 *
	 * @param slidePath - The ZIP path of the slide containing the graphic frame.
	 * @param graphicFrame - The raw XML object for the `p:graphicFrame` element.
	 * @returns The parsed chart data, or `undefined` if the frame is not a chart.
	 */
	public async getChartDataForGraphicFrame(
		slidePath: string,
		graphicFrame: XmlObject | undefined,
	): Promise<PptxChartData | undefined> {
		const graphicData = this.xmlLookupService.getChildByLocalName(
			this.xmlLookupService.getChildByLocalName(graphicFrame, 'graphic'),
			'graphicData',
		);
		const chartReference = this.xmlLookupService.getChildByLocalName(graphicData, 'chart');
		const chartRelationshipId = String(chartReference?.['@_r:id'] || '').trim();
		if (chartRelationshipId.length === 0) {
			return undefined;
		}

		const chartPart = await this.readXmlPartByRelationshipId(slidePath, chartRelationshipId);
		if (!chartPart) {
			return undefined;
		}

		const chartSpace = this.xmlLookupService.getChildByLocalName(chartPart.xml, 'chartSpace');
		const chartRoot = this.xmlLookupService.getChildByLocalName(chartSpace, 'chart');
		const plotArea = this.xmlLookupService.getChildByLocalName(chartRoot, 'plotArea');
		if (!plotArea) {
			return undefined;
		}

		// `c:clrMapOvr` remaps the 12 scheme-colour aliases (bg1/tx1/accent1...) for every
		// `a:schemeClr` resolved while parsing THIS chart's own XML, exactly as a slide's
		// `p:clrMapOvr` does for shape colours. Applied for the whole parse (classic and cx: charts
		// share the colour resolver) and restored in `finally` so it never leaks into later shapes.
		const clrMapOvr = this.parseClrMapOvr(chartSpace);
		const previousClrMapOverride = this.currentSlideClrMapOverride;
		if (clrMapOvr) {
			this.currentSlideClrMapOverride = clrMapOvr;
		}
		try {
			const { chartSpace: neutral, repairIssue } = readChartPartModel(chartPart.text);
			if (repairIssue) {
				this.compatibilityService.reportWarning({
					code: repairIssue.code,
					message: repairIssue.message,
					severity: 'info',
					scope: 'element',
					xmlPath: chartPart.partPath,
				});
			}
			// ChartEx (Office 2016+) parts use plotAreaRegion instead of chart groups.
			if (!neutral || neutral.plotArea.groups.length === 0) {
				return await this.parseCxChart(
					plotArea,
					this.detectChartType(plotArea),
					chartSpace,
					chartRoot,
					chartPart.partPath,
					chartRelationshipId,
				);
			}
			const chart = await this.buildClassicChart(
				neutral,
				{ chartSpace, chartRoot, plotArea },
				chartPart.partPath,
				chartRelationshipId,
			);
			return chart && (clrMapOvr ? { ...chart, clrMapOvr } : chart);
		} finally {
			this.currentSlideClrMapOverride = previousClrMapOverride;
		}
	}

	/** The runtime services the object-tree half of the adapter reads with. */
	private chartTreeReader(): ChartTreeReader {
		return {
			xml: this.xmlLookupService,
			parseColor: (node, placeholder) => this.parseColor(node, placeholder),
			resolveTypeface: (raw) => this.resolveThemeTypeface(raw) ?? raw,
			localName: (key) => this.compatibilityService.getXmlLocalName(key),
			...(this.colorStyleCodec ? { colorStyleCodec: this.colorStyleCodec } : {}),
			extractPointValues: (node, preferNumeric) =>
				this.extractChartPointValues(node, preferNumeric),
		};
	}

	/**
	 * Every series of every chart group, built from the neutral model with the styles of its
	 * object-tree node, plus the c15 filtered series and titles of the groups.
	 */
	private buildChartSeries(
		groups: ReturnType<typeof pairGroupNodes>,
		chartType: PptxChartData['chartType'],
		categories: string[],
		axes: NonNullable<PptxChartData['axes']>,
		reader: ChartTreeReader,
	) {
		const seriesNodes: XmlObject[] = [];
		const series: PptxChartData['series'] = [];
		const filtered: Pick<
			PptxChartData,
			'filteredSeries' | 'filteredSeriesTitle' | 'filteredCategoryTitle'
		> = {};
		for (const { group, node } of groups) {
			const hidden = parseFilteredSeries(node, reader.xml);
			if (hidden) filtered.filteredSeries = [...(filtered.filteredSeries ?? []), ...hidden];
			const titles = parseFilteredTitles(node, reader.xml);
			if (titles?.seriesTitle) filtered.filteredSeriesTitle ??= titles.seriesTitle;
			if (titles?.categoryTitle) filtered.filteredCategoryTitle ??= titles.categoryTitle;
			const containerChartType = groupChartType(group);
			// The value axis: searched from the end of the group's axis ids (scatter-like groups
			// list it last; category charts reference one value axis).
			const axisId = [...group.axisIds]
				.reverse()
				.find((id) => axes.some((axis) => axis.axisType === 'valAx' && axis.axisId === id));
			const nodes = rawChildren(node, 'ser', reader.localName);
			group.series.forEach((entry, index) => {
				const serNode = nodes[index] ?? {};
				seriesNodes.push(serNode);
				series.push(
					seriesFromNeutral(
						{
							series: entry,
							node: serNode,
							index,
							categories,
							...(chartType === 'combo' && containerChartType
								? { seriesChartType: containerChartType }
								: {}),
							...(containerChartType ? { containerChartType } : {}),
							...(axisId !== undefined ? { axisId } : {}),
						},
						reader,
					),
				);
			});
		}
		return { series, seriesNodes, filtered };
	}

	/**
	 * The drop lines, high-low lines and up/down bars, and the data table and 3-D surfaces, with
	 * their resolved styles. The line elements are legal only on some groups; on a combo whose
	 * first group is another type (volume + stock writes c:barChart first), each is read from the
	 * first group that allows it.
	 */
	private readChartLinesAndChrome(
		neutral: ChartSpace,
		tree: ChartTree,
		groups: ReturnType<typeof pairGroupNodes>,
		reader: ChartTreeReader,
	) {
		const colorAdapter = { parseColor: reader.parseColor };
		const container = (child: string): XmlObject | undefined =>
			(groups.find(({ group }) => chartContainerAllows(group.element, child)) ?? groups[0])?.node;
		const dropLines = parseLineStyle(container('dropLines'), 'dropLines', reader.xml, colorAdapter);
		const hiLowLines = parseLineStyle(
			container('hiLowLines'),
			'hiLowLines',
			reader.xml,
			colorAdapter,
		);
		const upDownBars = parseChartUpDownBars(container('upDownBars'), reader.xml, colorAdapter);
		const dataTable = neutral.plotArea.dataTable
			? parseDataTable(tree.plotArea, reader.xml, colorAdapter, reader.resolveTypeface)
			: undefined;
		const surfaces =
			tree.chartRoot && (neutral.floor || neutral.sideWall || neutral.backWall)
				? parseChart3DSurfaces(tree.chartRoot, reader.xml, colorAdapter)
				: {};
		return {
			...(dataTable ? { dataTable } : {}),
			...(dropLines ? { dropLines } : {}),
			...(hiLowLines ? { hiLowLines } : {}),
			...(upDownBars ? { upDownBars } : {}),
			...(surfaces.floor ? { floor: surfaces.floor } : {}),
			...(surfaces.sideWall ? { sideWall: surfaces.sideWall } : {}),
			...(surfaces.backWall ? { backWall: surfaces.backWall } : {}),
		};
	}

	/**
	 * A classic chart: structure from the neutral model, styles and the part facts from the
	 * object tree, the embedded workbook as the fallback of empty caches.
	 */
	private async buildClassicChart(
		neutral: ChartSpace,
		tree: ChartTree,
		chartPartPath: string,
		chartRelationshipId: string,
	): Promise<PptxChartData | undefined> {
		const { chartSpace, chartRoot, plotArea } = tree;
		const reader = this.chartTreeReader();
		const colorAdapter = { parseColor: reader.parseColor };
		const chartType = chartTypeFromNeutral(neutral);
		const groups = pairGroupNodes(neutral.plotArea.groups, plotArea, reader.localName);
		const axes = axesFromNeutral(
			neutral.plotArea.axes,
			parseChartAxes(plotArea, reader.xml, colorAdapter, reader.localName, reader.resolveTypeface),
		);
		const { categories, categoryLevels } = neutralCategories(neutral.plotArea.groups);
		const built = this.buildChartSeries(groups, chartType, categories, axes, reader);
		if (built.series.length === 0) {
			return undefined;
		}
		const firstNode = groups[0]?.node;
		const firstSeries = neutral.plotArea.groups.find((group) => group.series.length > 0)?.series[0];
		const rawDateCategories = axes.some((axis) => axis.axisType === 'dateAx')
			? neutralDateCategories(firstSeries?.categories)
			: undefined;
		// Lossless multi-run alternative to `title`, with the run styles resolved.
		const titleRuns = parseChartTitleRuns(
			this.xmlLookupService.getChildByLocalName(chartRoot, 'title'),
			reader.xml,
			colorAdapter,
		);
		const style = chartStyleFromNeutral(neutral, this.extractChartStyle(chartSpace, chartRoot));
		const groupOptions = groupOptionsFromNeutral(groups[0]?.group, chartType);
		const linesAndChrome = this.readChartLinesAndChrome(neutral, tree, groups, reader);

		const facts = await this.readChartPartFacts(chartSpace, chartRoot, chartPartPath);
		const workbook = facts.embeddedWorkbookData;
		// The chart's own `c:date1904` is authoritative over the embedded workbook's setting.
		const dateCategories = rawDateCategories
			? { ...rawDateCategories, date1904: neutral.date1904 ?? workbook?.date1904 ?? false }
			: undefined;
		const filled = applyEmbeddedWorkbookFallback(chartType, categories, built.series, workbook);

		// Resolve c:dPt/c:ser picture fills to image URLs (only when a series or point has one).
		const hasPicture = filled.series.some(
			(entry) =>
				entry.picture ||
				entry.impliedPicture ||
				entry.dataPoints?.some((point) => point.picture || point.impliedPicture),
		);
		if (hasPicture) {
			await resolveDataPointPictureImages(
				reader.xml,
				this.readChartRels.bind(this),
				this.resolveImagePath.bind(this),
				// A `data:` URL, which the shared renderer can decode synchronously (see
				// `getImageDataAsDataUrl`).
				this.getImageDataAsDataUrl.bind(this),
				built.seriesNodes,
				filled.series,
				chartPartPath,
			);
		}

		const ofPieOptions = chartType === 'ofPie' ? this.parseOfPieOptions(firstNode) : undefined;
		const bubbleOptions =
			chartType === 'bubble' ? parseBubbleChartOptions(firstNode, reader.localName) : undefined;
		const bandFmts =
			chartType === 'surface' && firstNode
				? parseChartBandFmts(firstNode, reader.xml, colorAdapter)
				: undefined;
		const view3D = view3DFromNeutral(neutral);
		const chartChrome = chartChromeFromNeutral(neutral);

		return {
			chartType,
			categories: filled.categories,
			...(categoryLevels ? { categoryLevels } : {}),
			...(dateCategories ? { dateCategories } : {}),
			series: filled.series,
			...built.filtered,
			title: neutralTitleText(neutral.title),
			...(titleRuns ? { titleRuns } : {}),
			style,
			grouping: groupOptions.grouping,
			...groupOptions,
			chartPartPath,
			chartRelationshipId,
			...linesAndChrome,
			...(axes.length > 0 ? { axes } : {}),
			...(bandFmts ? { bandFmts } : {}),
			...facts,
			...(neutral.plotVisibleOnly !== undefined
				? { plotVisibleOnly: neutral.plotVisibleOnly }
				: {}),
			...(ofPieOptions ? { ofPieOptions } : {}),
			...(bubbleOptions ? { bubbleOptions } : {}),
			...(view3D ? { view3D } : {}),
			...(chartChrome ? { chartChrome } : {}),
			...(neutral.date1904 !== undefined ? { date1904: neutral.date1904 } : {}),
			...(neutral.roundedCorners !== undefined ? { roundedCorners: neutral.roundedCorners } : {}),
		};
	}
}
