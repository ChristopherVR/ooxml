import {
	chartDrawingColorFromHex,
	chartDrawingGradientFrom,
	chartEffectEditsFromAppearance,
	chartPointsToEmu,
	chartThreeDFromAppearance,
	type ChartEditIntent,
	type ChartFormatTarget,
	type ChartEffectEdit,
	type ChartTextEdit,
} from 'ooxml-core/chart';
import type { ChartFormatElementView } from './format-pane-view';
import type { ChartFormatFieldValue } from './format-pane-types';
import { DEFAULT_CHART_GLOW, DEFAULT_CHART_SHADOW } from './format-pane-effects';
import { defaultPaneGradient, gradientEdit } from './format-pane-intent';
import { chartAnalysisFieldIntent } from './format-pane-analysis';

const num = (value: ChartFormatFieldValue): number | null =>
	value === '' || typeof value === 'boolean' || !Number.isFinite(Number(value))
		? null
		: Number(value);

function effectIntent(
	view: ChartFormatElementView,
	target: ChartFormatTarget,
	field: string,
	value: ChartFormatFieldValue,
): ChartEditIntent | null | undefined {
	const text = field.startsWith('text.');
	const name = text ? field.slice(5) : field;
	const match = /^(shadow|glow)\.(enabled|color|transparency|blur|distance|angle|radius)$/u.exec(
		name,
	);
	const wrap = (edit: ChartEffectEdit): ChartEditIntent =>
		text ? { op: 'text', target, edit: { effect: edit } } : { op: 'effect', target, edit };
	if (name === 'softEdge.size' && !text) {
		const n = num(value);
		return n === null || n < 0 || n > 100
			? null
			: wrap(
					n === 0
						? { kind: 'softEdge', remove: true }
						: { kind: 'softEdge', radiusEmu: chartPointsToEmu(n) },
				);
	}
	if (!match) return undefined;
	const kind = match[1] === 'shadow' ? 'outerShadow' : 'glow';
	const property = match[2]!;
	const effects = text ? view.text?.effects : view.effects;
	const current = match[1] === 'shadow' ? effects?.shadow : effects?.glow;
	if (property === 'enabled') {
		if (!value) return wrap({ kind, remove: true });
		return wrap(
			chartEffectEditsFromAppearance(
				kind === 'outerShadow' ? { shadow: DEFAULT_CHART_SHADOW } : { glow: DEFAULT_CHART_GLOW },
			)[0]!,
		);
	}
	if (property === 'color') {
		const color = chartDrawingColorFromHex(String(value), current?.transparency ?? 0);
		return color ? wrap(kind === 'outerShadow' ? { kind, color } : { kind: 'glow', color }) : null;
	}
	const n = num(value);
	if (n === null || n < 0) return null;
	if (property === 'transparency') return n <= 100 ? wrap({ kind, transparency: n }) : null;
	if (kind === 'glow')
		return property === 'radius' && n <= 150
			? wrap({ kind, radiusEmu: chartPointsToEmu(n) })
			: null;
	if (property === 'angle')
		return n <= 359 ? wrap({ kind, directionAngle: Math.round(n * 60000) }) : null;
	if (property === 'blur')
		return n <= 100 ? wrap({ kind, blurRadiusEmu: chartPointsToEmu(n) }) : null;
	return property === 'distance' && n <= 200
		? wrap({ kind, distanceEmu: chartPointsToEmu(n) })
		: null;
}

function threeDIntent(
	view: ChartFormatElementView,
	target: ChartFormatTarget,
	field: string,
	value: ChartFormatFieldValue,
): ChartEditIntent | null {
	const properties = view.threeD ?? {};
	const text = String(value);
	const n = num(value);
	const bevel = /^threeD\.bevel(Top|Bottom)(Width|Height)?$/u.exec(field);
	if (bevel) {
		const key = bevel[1] === 'Top' ? 'bevelTop' : 'bevelBottom';
		const current = properties[key] ?? { type: 'circle', width: 6, height: 6 };
		if (!bevel[2])
			return {
				op: 'threeD',
				target,
				edit: chartThreeDFromAppearance({
					[key]: text === 'none' ? null : { ...current, type: text },
				}),
			};
		if (n === null || n < 0 || n > 1584) return null;
		return {
			op: 'threeD',
			target,
			edit: chartThreeDFromAppearance({
				[key]: { ...current, [bevel[2] === 'Width' ? 'width' : 'height']: n },
			}),
		};
	}
	if (field === 'threeD.material' || field === 'threeD.lighting')
		return {
			op: 'threeD',
			target,
			edit: chartThreeDFromAppearance({
				[field === 'threeD.material' ? 'material' : 'lighting']: text,
			}),
		};
	const size = /^threeD\.(depth|contour)(Color|Size)$/u.exec(field);
	if (!size) return null;
	const key = size[1] === 'depth' ? 'depth' : 'contour';
	const current = properties[key] ?? { size: 0 };
	if (size[2] === 'Color')
		return chartDrawingColorFromHex(text)
			? {
					op: 'threeD',
					target,
					edit: chartThreeDFromAppearance({ [key]: { ...current, color: text } }),
				}
			: null;
	return n === null || n < 0 || n > 1584
		? null
		: { op: 'threeD', target, edit: chartThreeDFromAppearance({ [key]: { ...current, size: n } }) };
}

function textIntent(
	view: ChartFormatElementView,
	target: ChartFormatTarget,
	field: string,
	value: ChartFormatFieldValue,
): ChartEditIntent | null {
	const text = String(value);
	const n = num(value);
	const properties = view.text ?? {};
	const wrap = (edit: ChartTextEdit): ChartEditIntent => ({ op: 'text', target, edit });
	if (field === 'text.rotation')
		return value === ''
			? wrap({ rotation: null })
			: n === null || n < -90 || n > 90
				? null
				: wrap({ rotation: n });
	if (field === 'text.anchor')
		return ['t', 'ctr', 'b'].includes(text) ? wrap({ anchor: text }) : null;
	if (field === 'text.direction')
		return ['horz', 'vert', 'vert270', 'wordArtVert'].includes(text)
			? wrap({ direction: text })
			: null;
	if (field === 'text.fill.transparency' || field === 'text.outline.transparency')
		return n === null || n < 0 || n > 100
			? null
			: wrap(
					field === 'text.fill.transparency' ? { fillTransparency: n } : { lineTransparency: n },
				);
	const fill = field.startsWith('text.fill.');
	const prefix = fill ? 'text.fill.' : 'text.outline.';
	const key = field.slice(prefix.length);
	if (!field.startsWith(prefix)) return null;
	const current = fill ? properties.fill : properties.outline;
	if (key === 'kind' && text === 'automatic') return wrap(fill ? { fill: null } : { line: null });
	if (key === 'kind' && text === 'none')
		return wrap(fill ? { fill: { kind: 'none' } } : { line: { fill: { kind: 'none' } } });
	if (key === 'color' || (key === 'kind' && text === 'solid')) {
		const color = chartDrawingColorFromHex(
			key === 'color' ? text : (current?.color ?? '#000000'),
			current?.transparency ?? 0,
		);
		return color
			? wrap(
					fill ? { fill: { kind: 'solid', color } } : { line: { fill: { kind: 'solid', color } } },
				)
			: null;
	}
	if (fill && key === 'kind' && text === 'gradient') {
		const gradient = chartDrawingGradientFrom(
			properties.fill?.gradient ?? defaultPaneGradient(properties.fill?.color),
		);
		return gradient ? wrap({ fill: gradient }) : null;
	}
	if (fill && /^gradient(Type|Start|End|Angle)$/u.test(key)) {
		const gradient = gradientEdit(
			properties.fill?.gradient ?? defaultPaneGradient(properties.fill?.color),
			key.slice(8),
			value,
		);
		const next = gradient && chartDrawingGradientFrom(gradient);
		return next ? wrap({ fill: next }) : null;
	}
	if (!fill && key === 'width')
		return n === null || n < 0 || n > 1584
			? null
			: wrap({ line: { widthEmu: chartPointsToEmu(n) } });
	return null;
}

/** Undefined means a basic field; null means an invalid or unsupported advanced request. */
export function chartAdvancedFieldIntent(
	view: ChartFormatElementView,
	target: ChartFormatTarget,
	field: string,
	value: ChartFormatFieldValue,
): ChartEditIntent | null | undefined {
	try {
		if (field.startsWith('text.') && !view.support.text) return null;
		if (field.startsWith('threeD.') && !view.support.threeD) return null;
		if (/^(shadow|glow|softEdge)\./u.test(field) && view.support.effects === false) return null;
		if (
			(field.startsWith('fill.pattern') || (field === 'fill.kind' && value === 'pattern')) &&
			!view.support.pattern
		)
			return null;
		if (
			/^line\.(compound|join|beginArrow|endArrow)$/u.test(field) &&
			view.support.lineDetails === false
		)
			return null;
		const analysis = chartAnalysisFieldIntent(view.series, target, field, value);
		if (analysis !== undefined) return analysis;
		const effect = effectIntent(view, target, field, value);
		if (effect !== undefined) return effect;
		if (field.startsWith('threeD.')) return threeDIntent(view, target, field, value);
		if (field.startsWith('text.')) return textIntent(view, target, field, value);
		if (field === 'fill.kind' && value === 'pattern') return { op: 'pattern', target, edit: {} };
		if (field.startsWith('fill.pattern')) {
			const text = String(value);
			if (field === 'fill.patternPreset')
				return /^[A-Za-z][A-Za-z0-9]*$/u.test(text)
					? { op: 'pattern', target, edit: { preset: text } }
					: null;
			const color = chartDrawingColorFromHex(text);
			if (!color) return null;
			if (field === 'fill.patternForeground')
				return { op: 'pattern', target, edit: { foreground: color } };
			if (field === 'fill.patternBackground')
				return { op: 'pattern', target, edit: { background: color } };
			return null;
		}
		if (field === 'line.width' && value === '')
			return { op: 'lineOptions', target, edit: { widthEmu: null } };
		if (field === 'line.compound')
			return ['sng', 'dbl', 'thickThin', 'thinThick', 'tri'].includes(String(value))
				? { op: 'lineOptions', target, edit: { compound: String(value) } }
				: null;
		if (field === 'line.join')
			return ['round', 'bevel', 'miter'].includes(String(value))
				? { op: 'lineOptions', target, edit: { join: { kind: String(value) as 'round' } } }
				: null;
		if (field === 'line.beginArrow' || field === 'line.endArrow')
			return ['none', 'arrow', 'triangle', 'stealth', 'diamond', 'oval'].includes(String(value))
				? {
						op: 'lineOptions',
						target,
						edit:
							field === 'line.beginArrow'
								? { headEnd: { type: String(value) } }
								: { tailEnd: { type: String(value) } },
					}
				: null;
		if (field === 'series.axis')
			return target.kind === 'series' &&
				view.support.secondaryAxis &&
				['primary', 'secondary'].includes(String(value))
				? { op: 'secondaryAxis', seriesIndex: target.seriesIndex, secondary: value === 'secondary' }
				: null;
		return undefined;
	} catch {
		return null;
	}
}
