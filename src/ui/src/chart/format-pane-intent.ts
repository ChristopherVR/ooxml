/**
 * Turns one field change of a chart Format pane into a typed chart edit
 * intent (`ChartEditIntent` of `ooxml-core/chart`), the request every product
 * understands: `applyChartEditIntent` runs it over a `ChartSpace`, other
 * products map it onto their own edit path. Returns `null` for a value that is
 * not usable (an empty or out-of-range number), a target does not support it,
 * or the field requires host-owned media allocation. Advanced properties use
 * the same strict core operations as the product save adapters.
 *
 * @module chart/format-pane-intent
 */
import type {
	ChartAxisEdit,
	ChartEditIntent,
	ChartFormatTarget,
	ChartGradientFill,
} from 'ooxml-core/chart';
import type { DrawingColor, DrawingFill } from 'ooxml-core/drawingml';

import type { ChartFormatFieldValue } from './format-pane-types';
import type { ChartFormatElementView } from './format-pane-view';
import { chartAdvancedFieldIntent } from './format-pane-advanced-intent';
import { chartDrawingColorFromHex, chartDrawingGradientFrom } from 'ooxml-core/chart';

const number = (value: ChartFormatFieldValue): number | null => {
	if (value === '' || typeof value === 'boolean') return null;
	const n = typeof value === 'number' ? value : Number(value);
	return Number.isFinite(n) ? n : null;
};

/** `#RRGGBB` as a DrawingML colour, with the transparency (0..100) as `alpha`. */
export function drawingColorFromHex(hex: string, transparency = 0): DrawingColor | null {
	return chartDrawingColorFromHex(hex, transparency);
}

/** A pane gradient (hex stops, percent positions) as a DrawingML gradient fill. */
export function drawingGradientFrom(gradient: ChartGradientFill): DrawingFill | null {
	return chartDrawingGradientFrom(gradient);
}

/** Office's two-stop gradient from `color` towards white, what picking Gradient fill starts with. */
export function defaultPaneGradient(color = '#4472C4'): ChartGradientFill {
	return {
		type: 'linear',
		angle: 90,
		stops: [
			{ color, position: 0 },
			{ color: '#FFFFFF', position: 100 },
		],
	};
}

const GRADIENT = /^(fill|line)\.gradient(Type|Start|End|Angle)$/u;

export function gradientEdit(
	current: ChartGradientFill,
	part: string,
	value: ChartFormatFieldValue,
): ChartGradientFill | null {
	const stops = current.stops.map((stop) => ({ ...stop }));
	switch (part) {
		case 'Type':
			return value === 'linear' || value === 'radial' ? { ...current, type: value, stops } : null;
		case 'Start':
		case 'End': {
			const stop = part === 'Start' ? stops[0] : stops[stops.length - 1];
			if (!stop || !drawingColorFromHex(String(value))) return null;
			stop.color = String(value);
			return { ...current, stops };
		}
		default: {
			const n = number(value);
			return n === null ? null : { ...current, stops, angle: ((n % 360) + 360) % 360 };
		}
	}
}

function paintIntent(
	view: ChartFormatElementView,
	target: ChartFormatTarget,
	field: string,
	value: ChartFormatFieldValue,
): ChartEditIntent | null {
	const n = number(value);
	const text = String(value);
	const gradient = GRADIENT.exec(field);
	if (gradient) {
		const paint = gradient[1] === 'fill' ? view.fill : view.line;
		const next = gradientEdit(
			paint.gradient ?? defaultPaneGradient(paint.color),
			gradient[2]!,
			value,
		);
		const fill = next && drawingGradientFrom(next);
		if (!fill) return null;
		return gradient[1] === 'fill'
			? { op: 'fill', target, fill }
			: { op: 'line', target, line: { fill }, merge: true };
	}
	switch (field) {
		case 'fill.kind': {
			if (text === 'none') return { op: 'fill', target, fill: { kind: 'none' } };
			if (text === 'automatic') return { op: 'fill', target, fill: undefined };
			if (text === 'solid') {
				const color = drawingColorFromHex(view.fill.color ?? '#4472C4', view.fill.transparency);
				return color && { op: 'fill', target, fill: { kind: 'solid', color } };
			}
			if (text === 'gradient') {
				const fill = drawingGradientFrom(
					view.fill.gradient ?? defaultPaneGradient(view.fill.color),
				);
				return fill && { op: 'fill', target, fill };
			}
			return null;
		}
		case 'fill.color': {
			const color = drawingColorFromHex(text, view.fill.transparency);
			return color && { op: 'fill', target, fill: { kind: 'solid', color } };
		}
		case 'fill.transparency':
		case 'line.transparency':
			return n === null || n < 0 || n > 100
				? null
				: {
						op: 'transparency',
						target,
						paint: field === 'fill.transparency' ? 'fill' : 'line',
						percent: n,
					};
		case 'line.kind': {
			if (text === 'none')
				return { op: 'line', target, line: { fill: { kind: 'none' } }, merge: true };
			if (text === 'automatic') return { op: 'line', target, line: undefined };
			if (text === 'solid') {
				const color = drawingColorFromHex(view.line.color ?? '#000000', view.line.transparency);
				return (
					color && { op: 'line', target, line: { fill: { kind: 'solid', color } }, merge: true }
				);
			}
			if (text === 'gradient') {
				const fill = drawingGradientFrom(
					view.line.gradient ?? defaultPaneGradient(view.line.color),
				);
				return fill && { op: 'line', target, line: { fill }, merge: true };
			}
			return null;
		}
		case 'line.color': {
			const color = drawingColorFromHex(text, view.line.transparency);
			return color && { op: 'line', target, line: { fill: { kind: 'solid', color } }, merge: true };
		}
		case 'line.width':
			// Automatic width needs the outline without its width: no neutral operation drops it yet.
			return n === null || n < 0 || n > 1584
				? null
				: { op: 'line', target, line: { widthEmu: Math.round(n * 12700) }, merge: true };
		case 'line.dash':
			return { op: 'line', target, line: { dash: text }, merge: true };
		case 'line.cap':
			return { op: 'line', target, line: { cap: text }, merge: true };
		default:
			return null;
	}
}

const DISPLAY_UNITS = new Set([
	'hundreds',
	'thousands',
	'tenThousands',
	'hundredThousands',
	'millions',
	'tenMillions',
	'hundredMillions',
	'billions',
	'trillions',
]);

const one = <K extends keyof ChartAxisEdit>(key: K, value: ChartAxisEdit[K]): ChartAxisEdit => {
	const edit: ChartAxisEdit = {};
	edit[key] = value;
	return edit;
};

function axisEdit(field: string, value: ChartFormatFieldValue): ChartAxisEdit | null {
	const n = number(value);
	const text = String(value);
	switch (field) {
		case 'axis.title':
			return { title: text };
		case 'axis.numberFormatSourceLinked':
			return { numberFormatSourceLinked: Boolean(value) };
		case 'axis.baseTimeUnit':
		case 'axis.majorTimeUnit':
		case 'axis.minorTimeUnit': {
			const key = field.slice(5) as 'baseTimeUnit' | 'majorTimeUnit' | 'minorTimeUnit';
			return text === ''
				? { [key]: null }
				: ['days', 'months', 'years'].includes(text)
					? { [key]: text }
					: null;
		}
		case 'axis.min':
		case 'axis.max': {
			const key = field === 'axis.min' ? 'min' : 'max';
			return value === '' ? one(key, null) : n === null ? null : one(key, n);
		}
		case 'axis.majorUnit':
		case 'axis.minorUnit': {
			const key = field === 'axis.majorUnit' ? 'majorUnit' : 'minorUnit';
			if (value === '') return one(key, null);
			return n === null || n <= 0 ? null : one(key, n);
		}
		case 'axis.crosses':
			return text === 'autoZero' || text === 'min' || text === 'max'
				? { crosses: text }
				: text === 'value'
					? { crosses: 0 }
					: null;
		case 'axis.crossesAt':
			return n === null ? null : { crosses: n };
		case 'axis.displayUnits':
			return text === 'none'
				? { displayUnits: null }
				: DISPLAY_UNITS.has(text)
					? { displayUnits: text }
					: null;
		case 'axis.logScale':
			return { logBase: value ? 10 : null };
		case 'axis.logBase':
			return n === null || n < 2 || n > 1000 ? null : { logBase: n };
		case 'axis.reverseOrder':
			return { reverse: Boolean(value) };
		case 'axis.majorTickMark':
		case 'axis.minorTickMark':
			return text === 'none' || text === 'in' || text === 'out' || text === 'cross'
				? one(field === 'axis.majorTickMark' ? 'majorTickMark' : 'minorTickMark', text)
				: null;
		case 'axis.tickLabelPosition':
			return text === 'nextTo' || text === 'high' || text === 'low' || text === 'none'
				? { tickLabelPosition: text }
				: null;
		case 'axis.numberFormat':
			return { numberFormat: text.trim() };
		default:
			return null;
	}
}

type LabelsEdit = Extract<ChartEditIntent, { op: 'dataLabels' }>['edit'];
type LabelSwitch =
	| 'showSeriesName'
	| 'showCategoryName'
	| 'showValue'
	| 'showPercent'
	| 'showLegendKey';
const LABEL_SWITCHES: Partial<Record<string, LabelSwitch>> = {
	'labels.showSeriesName': 'showSeriesName',
	'labels.showCategory': 'showCategoryName',
	'labels.showValue': 'showValue',
	'labels.showPercent': 'showPercent',
	'labels.showLegendKey': 'showLegendKey',
};

const SPACING: Partial<Record<string, [keyof GroupEdit, number, number]>> = {
	'series.gapWidth': ['gapWidth', 0, 500],
	'series.overlap': ['overlap', -100, 100],
	'series.firstSliceAngle': ['firstSliceAngle', 0, 360],
	'series.holeSize': ['holeSize', 10, 90],
};
type GroupEdit = Extract<ChartEditIntent, { op: 'groupOptions' }>['edit'];

/** The intent of field `field` set to `value` on `target`, or `null` (see the module notes). */
export function chartFormatFieldIntent(
	view: ChartFormatElementView,
	target: ChartFormatTarget,
	field: string,
	value: ChartFormatFieldValue,
): ChartEditIntent | null {
	const advanced = chartAdvancedFieldIntent(view, target, field, value);
	if (advanced !== undefined) return advanced;
	if (field.startsWith('fill.') || field.startsWith('line.'))
		return paintIntent(view, target, field, value);
	if (field.startsWith('axis.')) {
		if (target.kind !== 'axis') return null;
		const edit = axisEdit(field, value);
		return edit && { op: 'axis', axisIndex: target.axisIndex, edit };
	}
	if (field === 'legend.position')
		return ['t', 'b', 'l', 'r', 'tr'].includes(String(value))
			? { op: 'legend', edit: { position: String(value) as 't' } }
			: null;
	if (field === 'legend.noOverlap') return { op: 'legend', edit: { overlay: !value } };
	if (field.startsWith('labels.') && target.kind === 'dataLabels') {
		const seriesIndex = target.seriesIndex;
		const scope = seriesIndex === undefined ? {} : { seriesIndex };
		const key = LABEL_SWITCHES[field];
		if (key) {
			const edit: LabelsEdit = {};
			edit[key] = Boolean(value);
			return { op: 'dataLabels', ...scope, edit };
		}
		return field === 'labels.position' && value !== ''
			? { op: 'dataLabels', ...scope, edit: { position: String(value) } }
			: null;
	}
	const spacing = SPACING[field];
	if (spacing && target.kind === 'series') {
		const n = number(value);
		const [key, min, max] = spacing;
		if (n === null || !Number.isInteger(n) || n < min || n > max) return null;
		const edit: GroupEdit = {};
		edit[key] = n;
		return { op: 'groupOptions', seriesIndex: target.seriesIndex, edit };
	}
	return null;
}
