/**
 * Turns one field change of a chart Format pane into a new chart model, via
 * the core's chart format operations. Returns `null` when the value is not
 * usable (an empty or out-of-range number), so the pane can repaint the
 * current value instead of committing.
 *
 * @module render/chart-format/chart-format-pane-edit
 */
import type {
	ChartAxisOptionsEdit,
	ChartElementFormatPatch,
	PptxChartData,
	PptxChartElementLine,
	PptxChartFormatTargetKey,
	PptxChartLineEndType,
} from 'ooxml-core/pptx';
import {
	readChartElementFormat,
	resetChartElementFormat,
	withChartAxisOptions,
	withChartDataLabelOptions,
	withChartElementFormat,
	withChartLegendOptions,
	withChartSeriesOptions,
} from 'ooxml-core/pptx';

import { DEFAULT_CHART_GLOW, DEFAULT_CHART_SHADOW } from './chart-format-pane-effects';
import type { ChartFormatFieldValue } from './chart-format-pane-types';
import { resolveChartFormatTarget } from './chart-format-targets';

function number(value: ChartFormatFieldValue): number | null {
	if (value === '' || typeof value === 'boolean') {
		return null;
	}
	const n = typeof value === 'number' ? value : Number(value);
	return Number.isFinite(n) ? n : null;
}

const clamp = (n: number, min: number, max: number) => Math.min(max, Math.max(min, n));

function formatPatch(
	data: PptxChartData,
	targetKey: PptxChartFormatTargetKey,
	field: string,
	value: ChartFormatFieldValue,
): ChartElementFormatPatch | null {
	const { target } = resolveChartFormatTarget(data, targetKey);
	const view = readChartElementFormat(data, target);
	const n = number(value);
	const text = String(value);
	switch (field) {
		case 'fill.kind':
			return text === 'none' || text === 'solid' || text === 'gradient' || text === 'automatic'
				? {
						fill: {
							kind: text,
							...(text === 'solid' && !view.fill.color ? { color: '#4472C4' } : {}),
						},
					}
				: null;
		case 'fill.color':
			return { fill: { kind: 'solid', color: text } };
		case 'fill.transparency':
			return n === null ? null : { fill: { transparency: clamp(n, 0, 100) } };
		case 'fill.gradientStart':
		case 'fill.gradientEnd':
		case 'fill.gradientAngle': {
			const gradient = view.fill.gradient;
			if (!gradient) {
				return null;
			}
			const stops = gradient.stops.map((stop) => ({ ...stop }));
			if (field === 'fill.gradientStart' && stops[0]) {
				stops[0].color = text;
			} else if (field === 'fill.gradientEnd' && stops.length > 0) {
				stops[stops.length - 1]!.color = text;
			} else if (field === 'fill.gradientAngle' && n === null) {
				return null;
			}
			const angle =
				field === 'fill.gradientAngle' && n !== null ? clamp(n, 0, 359) : gradient.angle;
			return {
				fill: {
					kind: 'gradient',
					gradient: { ...gradient, stops, ...(angle !== undefined ? { angle } : {}) },
				},
			};
		}
		case 'line.kind':
			return text === 'none' || text === 'solid' || text === 'automatic'
				? {
						line: {
							kind: text,
							...(text === 'solid' && !view.line.color ? { color: '#000000' } : {}),
						},
					}
				: null;
		case 'line.color':
			return { line: { kind: 'solid', color: text } };
		case 'line.transparency':
			return n === null ? null : { line: { transparency: clamp(n, 0, 100) } };
		case 'line.width':
			return n === null ? null : { line: { width: clamp(n, 0, 1584) } };
		case 'line.dash':
			return { line: { dash: text } };
		case 'line.compound':
			return { line: { compound: text as PptxChartElementLine['compound'] } };
		case 'line.cap':
			return { line: { cap: text as PptxChartElementLine['cap'] } };
		case 'line.join':
			return { line: { join: text as PptxChartElementLine['join'] } };
		case 'line.beginArrow':
			return { line: { beginArrow: text as PptxChartLineEndType } };
		case 'line.endArrow':
			return { line: { endArrow: text as PptxChartLineEndType } };
		case 'shadow.enabled':
			return { effects: { shadow: value ? { ...DEFAULT_CHART_SHADOW } : null } };
		case 'glow.enabled':
			return { effects: { glow: value ? { ...DEFAULT_CHART_GLOW } : null } };
		case 'softEdge.size':
			return n === null ? null : { effects: { softEdge: n > 0 ? clamp(n, 0, 100) : null } };
		default: {
			const [effect, prop] = field.split('.');
			if (effect === 'shadow' && view.effects.shadow && prop) {
				const shadow = { ...view.effects.shadow };
				if (prop === 'color') {
					shadow.color = text;
				} else if (n === null) {
					return null;
				} else if (
					prop === 'transparency' ||
					prop === 'blur' ||
					prop === 'distance' ||
					prop === 'angle'
				) {
					shadow[prop] =
						prop === 'angle' ? clamp(n, 0, 359) : clamp(n, 0, prop === 'transparency' ? 100 : 200);
				}
				return { effects: { shadow } };
			}
			if (effect === 'glow' && view.effects.glow && prop) {
				const glow = { ...view.effects.glow };
				if (prop === 'color') {
					glow.color = text;
				} else if (n === null) {
					return null;
				} else if (prop === 'radius' || prop === 'transparency') {
					glow[prop] = clamp(n, 0, prop === 'transparency' ? 100 : 150);
				}
				return { effects: { glow } };
			}
			return null;
		}
	}
}

function axisEdit(field: string, value: ChartFormatFieldValue): ChartAxisOptionsEdit | null {
	const n = number(value);
	switch (field) {
		case 'axis.min':
		case 'axis.max':
		case 'axis.majorUnit':
		case 'axis.minorUnit': {
			const key = field.slice(5) as 'min' | 'max' | 'majorUnit' | 'minorUnit';
			if (value === '') {
				return { [key]: null };
			}
			if (n === null || ((key === 'majorUnit' || key === 'minorUnit') && n <= 0)) {
				return null;
			}
			return { [key]: n };
		}
		case 'axis.crosses':
			return value === 'value'
				? { crosses: 0 }
				: value === 'autoZero' || value === 'min' || value === 'max'
					? { crosses: value }
					: null;
		case 'axis.crossesAt':
			return n === null ? null : { crosses: n };
		case 'axis.numberFormat':
			return { numberFormat: String(value).trim() || null };
		case 'axis.majorTickMark':
		case 'axis.minorTickMark':
			return ['none', 'in', 'out', 'cross'].includes(String(value))
				? { [field.slice(5)]: value as 'none' }
				: null;
		case 'axis.tickLabelPosition':
			return { tickLabelPosition: value as 'nextTo' };
		case 'axis.displayUnits':
			return { displayUnits: value === 'none' ? null : (value as 'thousands') };
		case 'axis.logScale':
			return { logBase: value ? 10 : null };
		case 'axis.logBase':
			return n === null || n < 2 || n > 1000 ? null : { logBase: n };
		case 'axis.reverseOrder':
			return { reverseOrder: Boolean(value) };
		default:
			return null;
	}
}

/**
 * Apply one pane field change to `data` for the element `targetKey`, or
 * return `null` when the value cannot be applied.
 */
export function applyChartFormatField(
	data: PptxChartData,
	targetKey: PptxChartFormatTargetKey,
	field: string,
	value: ChartFormatFieldValue,
): PptxChartData | null {
	const { target } = resolveChartFormatTarget(data, targetKey);
	try {
		if (field.startsWith('series.')) {
			const n = number(value);
			if (n === null) {
				return null;
			}
			const key = field.slice(7) as 'gapWidth' | 'overlap' | 'firstSliceAngle' | 'holeSize';
			return withChartSeriesOptions(data, { [key]: Math.round(n) });
		}
		if (field.startsWith('axis.')) {
			const edit = target.kind === 'axis' ? axisEdit(field, value) : null;
			return edit && target.kind === 'axis'
				? withChartAxisOptions(data, target.axisIndex, edit)
				: null;
		}
		if (field === 'legend.position') {
			return withChartLegendOptions(data, { position: value as 't' });
		}
		if (field === 'legend.noOverlap') {
			return withChartLegendOptions(data, { overlay: !value });
		}
		if (field === 'labels.position') {
			return withChartDataLabelOptions(data, { position: value === '' ? null : (value as 'ctr') });
		}
		if (field.startsWith('labels.')) {
			return withChartDataLabelOptions(data, { [field.slice(7)]: Boolean(value) });
		}
		const patch = formatPatch(data, targetKey, field, value);
		return patch ? withChartElementFormat(data, target, patch) : null;
	} catch {
		// Out-of-range series spacing (the core validates against the schema).
		return null;
	}
}

/** Reset to Match Style for `targetKey`. */
export function resetChartFormatTarget(
	data: PptxChartData,
	targetKey: PptxChartFormatTargetKey,
): PptxChartData {
	return resetChartElementFormat(data, resolveChartFormatTarget(data, targetKey).target);
}
