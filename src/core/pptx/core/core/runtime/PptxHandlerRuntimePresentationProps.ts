import { XmlObject } from '../../types';
import type { PptxPresentationProperties, PptxChartStyle, PptxViewProperties } from '../../types';
import type { ChartAreaFormat } from '../../utils/chart-area-format';
import { parseChartAreaFormat } from '../../utils/chart-area-format';
import { parseChartDataLabelOptions } from '../../utils/chart-data-label-parser';
import {
	dataLabelsGroupDeleted,
	dataLabelsGroupShowsContent,
} from '../../utils/chart-data-labels-visibility';
import {
	parseChartWideTextStyle,
	parseDefRPrTextStyle,
	resolveTxPrDefRPr,
} from '../../utils/chart-def-rpr-style';
import { parseChartGradientFill } from '../../utils/chart-gradient-fill';
import { parseChartLegendEntries } from '../../utils/chart-legend-serializer';
import { parseChartTitleStyle } from '../../utils/chart-title-style-parser';
import { parseShowProperties } from './pptx-presentation-props-helpers';
import { findChildByLocalName, parsePrintProperties } from './pptx-print-properties';
import { parseViewProperties } from './pptx-view-props-helpers';
import { PptxHandlerRuntime as PptxHandlerRuntimeBase } from './PptxHandlerRuntimeSlideMasters';

export class PptxHandlerRuntime extends PptxHandlerRuntimeBase {
	/**
	 * Forward declaration - implemented in PptxHandlerRuntimeTextStyleUtils.
	 * Resolves `+mj-lt` / `+mn-lt` theme font tokens to the theme typeface.
	 */
	protected resolveThemeTypeface(_typeface: string | undefined): string | undefined {
		throw new Error('resolveThemeTypeface not yet initialised');
	}

	/**
	 * Parse presentation properties from `presentationPr.xml`.
	 * Extracts show type, loop, narration, animation, and print settings.
	 */
	protected async parsePresentationProperties(): Promise<PptxPresentationProperties | undefined> {
		try {
			// First find presentationPr relationship
			const relsXml = await this.zip.file('ppt/_rels/presentation.xml.rels')?.async('string');
			if (!relsXml) {
				return undefined;
			}

			const relsData = this.parser.parse(relsXml);
			const rels = this.ensureArray(relsData?.Relationships?.Relationship);
			const prRel = rels.find(
				(r: XmlObject) =>
					String(r?.['@_Type'] || '').includes('presProps') ||
					String(r?.['@_Target'] || '').includes('presProps'),
			);

			const prTarget = prRel ? String(prRel['@_Target'] || '') : 'presProps.xml';
			const prPath = prTarget.startsWith('/') ? prTarget.substring(1) : `ppt/${prTarget}`;

			const prXmlStr = await this.zip.file(prPath)?.async('string');
			if (!prXmlStr) {
				return undefined;
			}

			const prXml = this.parser.parse(prXmlStr);
			const rootKey = Object.keys(prXml ?? {}).find(
				(key) => this.compatibilityService.getXmlLocalName(key) === 'presentationPr',
			);
			const presProps = rootKey ? (prXml[rootKey] as XmlObject | undefined) : undefined;
			if (!presProps) {
				return undefined;
			}

			const props: PptxPresentationProperties = {};

			// Show properties (p:showPr)
			const showPr = presProps['p:showPr'] as XmlObject | undefined;
			if (showPr) {
				Object.assign(
					props,
					parseShowProperties(showPr, (node) => this.parseColor(node)),
				);
			}

			// Print properties (p:prnPr)
			const prnPr = findChildByLocalName(presProps, 'prnPr');
			if (prnPr) {
				props.printProperties = parsePrintProperties(prnPr);
			}

			// Most-recently-used colours (p:clrMru)
			const clrMru = presProps['p:clrMru'] as XmlObject | undefined;
			if (clrMru) {
				const colorNodes = this.ensureArray(clrMru['a:srgbClr']);
				const mruColors = colorNodes
					.map((c: XmlObject) => {
						const val = String(c?.['@_val'] || '').trim();
						return val.length > 0 ? `#${val}` : '';
					})
					.filter((c: string) => c.length > 0);
				if (mruColors.length > 0) {
					props.mruColors = mruColors;
				}
			}

			// NOTE: `p:gridSpacing` does NOT live under `p:presentationPr` in real
			// PowerPoint files; it lives under `p:viewPr` in `ppt/viewProps.xml`.
			// It used to be (incorrectly) read here, which meant this field was
			// always `undefined` for real decks. See `parseViewProperties` below
			// and `pptx-view-props-helpers.ts` for the correct read; consumers
			// must use `PptxData.viewProperties.gridSpacing`.

			return props;
		} catch (e) {
			console.warn('Failed to parse presentation properties:', e);
			return undefined;
		}
	}

	/**
	 * Parse view properties from `ppt/viewProps.xml`.
	 */
	protected async parseViewProperties(): Promise<PptxViewProperties | undefined> {
		try {
			const viewPropsXml = await this.zip.file('ppt/viewProps.xml')?.async('string');
			if (!viewPropsXml) {
				return undefined;
			}

			const data = this.parser.parse(viewPropsXml) as XmlObject;
			const rootKey = Object.keys(data ?? {}).find((key) => key.replace(/^.*:/u, '') === 'viewPr');
			const viewPrRoot = (rootKey ? data[rootKey] : undefined) as XmlObject | undefined;
			if (!viewPrRoot) {
				return undefined;
			}

			return parseViewProperties(viewPrRoot);
		} catch (e) {
			console.warn('Failed to parse view properties:', e);
			return undefined;
		}
	}

	/**
	 * Resolve the `c:spPr` fill and border of a chart container
	 * (`c:chartSpace` or `c:plotArea`) to a colour string, or the literal
	 * `'none'` for `<a:noFill/>`; `undefined` when the container declares
	 * nothing, which leaves the choice to the renderer. The save path reads it
	 * back with the same function (`chart-area-format.ts`).
	 */
	private parseChartContainerFormat(container: XmlObject | undefined): ChartAreaFormat {
		return parseChartAreaFormat(
			container,
			(key) => this.compatibilityService.getXmlLocalName(key),
			(node) => this.parseColor(node),
		);
	}

	/**
	 * Whether a legend overlays the plot. A classic `c:overlay` is a
	 * `CT_Boolean`, so `<c:overlay/>` with no `val` is true; a classic legend
	 * with no `c:overlay` at all also overlays (PowerPoint 16, every position).
	 * ChartEx (`cx:chartData` in the chart space) is only read from an explicit
	 * `@overlay`, since its missing-attribute behaviour has not been measured.
	 */
	private resolveLegendOverlay(
		chartSpace: XmlObject | undefined,
		legend: XmlObject,
	): boolean | undefined {
		const isFalse = (value: unknown) => value === '0' || value === 'false' || value === false;
		const overlay = this.xmlLookupService.getChildByLocalName(legend, 'overlay');
		if (overlay) {
			return !isFalse(overlay['@_val']);
		}
		if (legend['@_overlay'] !== undefined) {
			return !isFalse(legend['@_overlay']);
		}
		const isChartEx = this.xmlLookupService.getChildByLocalName(chartSpace, 'chartData');
		return isChartEx ? undefined : true;
	}

	/**
	 * Extract chart style metadata from chart XML.
	 */
	protected extractChartStyle(
		chartSpace: XmlObject | undefined,
		chartRoot: XmlObject | undefined,
	): PptxChartStyle | undefined {
		if (!chartSpace && !chartRoot) {
			return undefined;
		}
		const style: PptxChartStyle = {};
		let hasStyle = false;

		// Style ID from c:style. Office writes it inside `mc:AlternateContent` (`c14:style` under
		// `mc:Choice`, `c:style` under `mc:Fallback`); a consumer that does not model `c14:style`
		// takes the fallback (ECMA-376 Part 3, markup compatibility), as `parseChartSpace` does.
		const styleNode =
			this.xmlLookupService.getChildByLocalName(chartSpace, 'style') ??
			this.xmlLookupService
				.getChildrenArrayByLocalName(chartSpace, 'AlternateContent')
				.map((alternate) =>
					this.xmlLookupService.getChildByLocalName(
						this.xmlLookupService.getChildByLocalName(alternate, 'Fallback'),
						'style',
					),
				)
				.find((node) => node?.['@_val'] !== undefined);
		if (styleNode?.['@_val']) {
			style.styleId = parseInt(String(styleNode['@_val']));
			hasStyle = true;
		}

		// Chart-area fill (`c:chartSpace/c:spPr`). `<a:noFill/>` is the common
		// case and means the chart floats on the slide background; recording it as
		// `'none'` stops renderers painting their own panel behind it.
		const chartArea = this.parseChartContainerFormat(chartSpace);
		if (chartArea.fill) {
			style.chartAreaFill = chartArea.fill;
			hasStyle = true;
		}
		if (chartArea.border) {
			style.chartAreaBorder = chartArea.border;
			hasStyle = true;
		}
		const chartAreaGradient = parseChartGradientFill(
			this.xmlLookupService.getChildByLocalName(chartSpace, 'spPr'),
			this.xmlLookupService,
			this.colorStyleCodec,
		);
		if (chartAreaGradient) {
			style.chartAreaGradient = chartAreaGradient;
			hasStyle = true;
		}
		const textStyle = parseChartWideTextStyle(
			chartSpace,
			this.xmlLookupService,
			{ parseColor: (node) => this.parseColor(node) },
			(raw) => this.resolveThemeTypeface(raw) ?? raw,
		);
		if (textStyle) {
			style.textStyle = textStyle;
			hasStyle = true;
		}

		if (chartRoot) {
			// Legend
			const legend = this.xmlLookupService.getChildByLocalName(chartRoot, 'legend');
			if (legend) {
				style.hasLegend = true;
				hasStyle = true;
				// Classic charts nest position in a child (`c:legend/c:legendPos/@val`);
				// ChartEx (`cx:`) charts put it directly on the element
				// (`cx:legend/@pos`). Fall back to the attribute when the child lookup
				// misses so a `cx:legend pos="t"` isn't silently ignored (it used to
				// always fall through to renderers' `?? 'b'` default).
				const legendPos = this.xmlLookupService.getChildByLocalName(legend, 'legendPos');
				const legendPosVal = legendPos?.['@_val'] ?? legend['@_pos'];
				if (legendPosVal) {
					style.legendPosition = String(legendPosVal);
				}
				const legendOverlay = this.resolveLegendOverlay(chartSpace, legend);
				if (legendOverlay !== undefined) {
					style.legendOverlay = legendOverlay;
				}
				const entries = parseChartLegendEntries(
					legend,
					(key) => this.compatibilityService.getXmlLocalName(key),
					(node) => this.parseColor(node),
					(raw) => this.resolveThemeTypeface(raw) ?? raw,
				);
				if (entries.length > 0) {
					style.legendEntries = entries;
				}
				// The legend's own default text style (`c:legend/c:txPr`), falling
				// back to the chart-wide default (`c:chartSpace/c:txPr`) when the
				// legend declares none of its own. Every legend entry without a
				// per-entry `c:legendEntry/c:txPr` override renders at this size
				// instead of a fixed 9px default (issue: a chart authoring an 18pt
				// chart-level txPr rendered its legend at the hardcoded default).
				const legendTxPr = this.xmlLookupService.getChildByLocalName(legend, 'txPr');
				const chartSpaceTxPr = this.xmlLookupService.getChildByLocalName(chartSpace, 'txPr');
				const legendDefRPr = resolveTxPrDefRPr(legendTxPr ?? chartSpaceTxPr, this.xmlLookupService);
				const legendTextStyle = parseDefRPrTextStyle(
					legendDefRPr,
					this.xmlLookupService,
					{ parseColor: (node) => this.parseColor(node) },
					(raw) => this.resolveThemeTypeface(raw) ?? raw,
				);
				if (legendTextStyle) {
					style.legendTextStyle = legendTextStyle;
				}
			}

			// Title
			const title = this.xmlLookupService.getChildByLocalName(chartRoot, 'title');
			if (title) {
				style.hasTitle = true;
				hasStyle = true;
				Object.assign(
					style,
					parseChartTitleStyle(
						title,
						this.xmlLookupService,
						{ parseColor: (node, placeholder) => this.parseColor(node, placeholder) },
						(raw) => this.resolveThemeTypeface(raw) ?? raw,
					),
				);
			}

			// Plot area gridlines
			const plotArea = this.xmlLookupService.getChildByLocalName(chartRoot, 'plotArea');
			if (plotArea) {
				const plotAreaGradient = parseChartGradientFill(
					this.xmlLookupService.getChildByLocalName(plotArea, 'spPr'),
					this.xmlLookupService,
					this.colorStyleCodec,
				);
				if (plotAreaGradient) {
					style.plotAreaGradient = plotAreaGradient;
				}
				const plotAreaFormat = this.parseChartContainerFormat(plotArea);
				if (plotAreaFormat.fill) {
					style.plotAreaFill = plotAreaFormat.fill;
					hasStyle = true;
				}
				if (plotAreaFormat.border) {
					style.plotAreaBorder = plotAreaFormat.border;
					hasStyle = true;
				}
				const valAx = this.xmlLookupService.getChildByLocalName(plotArea, 'valAx');
				if (valAx) {
					const majorGridlines = this.xmlLookupService.getChildByLocalName(valAx, 'majorGridlines');
					if (majorGridlines) {
						style.hasGridlines = true;
						hasStyle = true;
					}
				}

				// Data labels check across chart types
				const chartTypeKeys = Object.keys(plotArea).filter((key) =>
					this.compatibilityService.getXmlLocalName(key).endsWith('Chart'),
				);
				for (const ctKey of chartTypeKeys) {
					const ctNode = plotArea[ctKey] as XmlObject | undefined;
					if (!ctNode) {
						continue;
					}

					// Check chart-level dLbls (applies to all series). PowerPoint
					// writes an all-zero group on every chart it authors, so the
					// group's presence alone is NOT "labels on": leave the flag
					// undefined (round-trips untouched) unless a show* flag is set.
					const chartDLbls = this.xmlLookupService.getChildByLocalName(ctNode, 'dLbls');
					if (chartDLbls && !style.dataLabels) {
						if (dataLabelsGroupDeleted(chartDLbls, this.xmlLookupService)) {
							style.hasDataLabels = false;
						} else if (dataLabelsGroupShowsContent(chartDLbls, this.xmlLookupService)) {
							style.hasDataLabels = true;
						}
						style.dataLabels = parseChartDataLabelOptions(
							chartDLbls,
							this.xmlLookupService,
							{ parseColor: (node, placeholder) => this.parseColor(node, placeholder) },
							(raw) => this.resolveThemeTypeface(raw) ?? raw,
						);
						hasStyle = true;
					}

					// Also check per-series dLbls
					if (!style.hasDataLabels) {
						const seriesList = this.xmlLookupService.getChildrenArrayByLocalName(ctNode, 'ser');
						for (const ser of seriesList) {
							const dLbls = this.xmlLookupService.getChildByLocalName(ser, 'dLbls');
							if (dataLabelsGroupShowsContent(dLbls, this.xmlLookupService)) {
								style.hasDataLabels = true;
								hasStyle = true;
							}
						}
					}
				}
			}
		}

		return hasStyle ? style : undefined;
	}

	protected toPresentationTarget(slidePath: string): string {
		const normalized = slidePath.startsWith('/') ? slidePath.substring(1) : slidePath;
		return normalized.startsWith('ppt/') ? normalized.substring(4) : normalized;
	}

	protected toSlidePathFromTarget(target: string): string {
		const normalized = target.startsWith('/') ? target.substring(1) : target;
		return normalized.startsWith('ppt/') ? normalized : `ppt/${normalized}`;
	}

	protected toSlideRelsPath(slidePath: string): string {
		return `${slidePath.replace('slides/', 'slides/_rels/')}.rels`;
	}
}
