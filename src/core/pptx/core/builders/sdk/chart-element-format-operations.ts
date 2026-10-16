/**
 * Headless operations behind PowerPoint's chart Format task panes (Fill,
 * Border/Line and Effects of one chart element). Each is pure: it returns a
 * new `PptxChartData` and leaves its input alone, so a UI can apply it as one
 * undoable element patch.
 *
 * An edit is recorded in `PptxChartData.elementFormats` (what the save path
 * writes into the element's `c:spPr`) and mirrored onto the flat model
 * fields the renderer already reads (`style.chartAreaFill`, a series'
 * `color` / `lineWidth`, an axis' `majorGridlinesSpPr`, ...), so the change
 * shows on the slide at once.
 *
 * @module sdk/chart-element-format-operations
 */
import type {
	PptxChartData,
	PptxChartElementEffects,
	PptxChartElementFill,
	PptxChartElementFormat,
	PptxChartElementLine,
	PptxChartElementText,
	PptxChartElementThreeD,
	PptxChartFormatTarget,
	PptxChartGradientFill,
	PptxChartPatternFill,
	PptxChartPictureFill,
} from '../../types';
import { chartFormatTargetKey } from '../../utils/chart-element-format-keys';
import {
	applyGradient,
	hasGradientField,
	mirrorFlatFields,
	withSeries,
} from './chart-element-format-mirror';
import { readChartElementFormat } from './chart-element-format-read';
import type { ChartFormatFillKind } from './chart-element-format-read';

export { readChartElementFormat, chartFormatTargetSupports } from './chart-element-format-read';
export type {
	ChartElementFormatView,
	ChartFormatFillKind,
	ChartFormatTargetSupport,
} from './chart-element-format-read';

/** What a pane edit changes; omitted sections and fields stay as they are. */
export interface ChartElementFormatPatch {
	fill?: Partial<Omit<PptxChartElementFill, 'kind' | 'pattern' | 'picture'>> & {
		kind?: ChartFormatFillKind;
		gradient?: PptxChartGradientFill;
		pattern?: Partial<PptxChartPatternFill>;
		/** A new `src` replaces the picture (and drops a loaded picture's relationship). */
		picture?: Partial<PptxChartPictureFill>;
	};
	line?: Partial<PptxChartElementLine>;
	effects?: PptxChartElementEffects;
	threeD?: PptxChartElementThreeD;
	/** Text Options; `fill` and `outline` merge field by field. */
	text?: Omit<PptxChartElementText, 'fill' | 'outline'> & {
		fill?: Partial<NonNullable<PptxChartElementText['fill']>>;
		outline?: Partial<NonNullable<PptxChartElementText['outline']>>;
	};
}

/**
 * Apply a Format-pane edit to one chart element.
 *
 * @example
 * ```ts
 * const next = withChartElementFormat(chart.chartData, { kind: 'legend' }, {
 *   fill: { kind: 'solid', color: '#FFF2CC', transparency: 20 },
 *   line: { kind: 'solid', color: '#BF9000', width: 1.5, dash: 'dash' },
 * });
 * ```
 */
export function withChartElementFormat(
	data: PptxChartData,
	target: PptxChartFormatTarget,
	patch: ChartElementFormatPatch,
): PptxChartData {
	const key = chartFormatTargetKey(target);
	const view = readChartElementFormat(data, target);
	const previous: PptxChartElementFormat = data.elementFormats?.[key] ?? {};
	const entry: PptxChartElementFormat = { ...previous };
	let next = data;
	let fill: PptxChartElementFill | undefined;
	if (patch.fill && view.support.fill) {
		const kind = patch.fill.kind ?? view.fill.kind;
		if (kind === 'gradient') {
			const gradient = patch.fill.gradient ?? view.fill.gradient;
			if (gradient && hasGradientField(target)) {
				// Chart area, plot area and series keep their gradient in the model's own field.
				next = applyGradient(next, target, gradient);
				delete entry.fill;
			} else if (gradient) {
				fill = { kind: 'gradient', gradient };
				entry.fill = fill;
			}
		} else if (kind === 'pattern') {
			const base = view.fill.pattern ?? {
				preset: 'pct5',
				foreground: '#4472C4',
				background: '#FFFFFF',
			};
			fill = { kind: 'pattern', pattern: { ...base, ...patch.fill.pattern } };
			entry.fill = fill;
		} else if (kind === 'picture') {
			const previous = view.fill.kind === 'picture' ? view.fill.picture : undefined;
			const picture: PptxChartPictureFill = { ...previous, ...patch.fill.picture };
			if (patch.fill.picture?.src && patch.fill.picture.src !== previous?.src) {
				delete picture.relId;
			}
			if (!picture.src && !picture.relId) {
				return data;
			}
			fill = { kind: 'picture', picture };
			entry.fill = fill;
		} else {
			fill = {
				kind,
				...((patch.fill.color ?? view.fill.color)
					? { color: patch.fill.color ?? view.fill.color }
					: {}),
				...((patch.fill.transparency ?? view.fill.transparency) !== undefined
					? { transparency: patch.fill.transparency ?? view.fill.transparency }
					: {}),
			};
			entry.fill = fill;
		}
	}
	let line: PptxChartElementLine | undefined;
	if (patch.line) {
		line = { ...view.line, ...patch.line };
		if (line.kind !== 'gradient') {
			delete line.gradient;
		}
		entry.line = line;
	}
	if (patch.effects) {
		entry.effects = { ...view.effects, ...patch.effects };
	}
	if (patch.threeD && view.support.threeD) {
		entry.threeD = { ...view.threeD, ...patch.threeD };
	}
	if (patch.text && view.support.text) {
		const text: PptxChartElementText = { ...view.text, ...patch.text } as PptxChartElementText;
		if (patch.text.fill) {
			text.fill = { kind: 'automatic', ...view.text.fill, ...patch.text.fill };
			if (text.fill.kind !== 'gradient') delete text.fill.gradient;
		}
		if (patch.text.outline) {
			text.outline = { kind: 'automatic', ...view.text.outline, ...patch.text.outline };
		}
		if (patch.text.effects) {
			text.effects = { ...view.text.effects, ...patch.text.effects };
		}
		entry.text = text;
	}
	next = mirrorFlatFields(next, target, fill, line);
	return { ...next, elementFormats: { ...next.elementFormats, [key]: entry } };
}

/**
 * PowerPoint's Reset to Match Style: drop the element's own formatting so the
 * chart style draws it again.
 */
export function resetChartElementFormat(
	data: PptxChartData,
	target: PptxChartFormatTarget,
): PptxChartData {
	const key = chartFormatTargetKey(target);
	let next: PptxChartData = data;
	switch (target.kind) {
		case 'chartArea':
		case 'plotArea': {
			const style = { ...data.style };
			delete style[`${target.kind}Fill`];
			delete style[`${target.kind}Gradient`];
			delete style[`${target.kind}Border`];
			next = { ...data, style };
			break;
		}
		case 'title':
			next = { ...data, style: { ...data.style, titleSpPr: null } };
			break;
		case 'axis':
		case 'majorGridlines':
		case 'minorGridlines': {
			const field =
				target.kind === 'axis'
					? 'spPr'
					: target.kind === 'majorGridlines'
						? 'majorGridlinesSpPr'
						: 'minorGridlinesSpPr';
			next = {
				...data,
				axes: (data.axes ?? []).map((axis, i) => {
					if (i !== target.axisIndex) {
						return axis;
					}
					const copy = { ...axis };
					delete copy[field];
					return copy;
				}),
			};
			break;
		}
		case 'series':
			next = withSeries(data, target.seriesIndex, (series) => {
				const copy = { ...series };
				delete copy.color;
				delete copy.gradientFill;
				delete copy.lineWidth;
				delete copy.lineDashStyle;
				delete copy.lineNoFill;
				return copy;
			});
			break;
		default:
			break;
	}
	return { ...next, elementFormats: { ...next.elementFormats, [key]: { reset: true } } };
}
