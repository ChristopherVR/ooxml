import type { DrawingColor, DrawingFill, DrawingLine } from '../../drawingml/types';
import { drawingFillXml } from '../../drawingml/write-fill';
import { NS, buildXml, first, parseXml } from '../../xml/index';
import type { ChartGradientFill } from '../gradient-definition';
import type {
	ChartEffectsAppearance,
	ChartFillAppearance,
	ChartLineAppearance,
	ChartTextAppearance,
	ChartThreeDAppearance,
} from './appearance-types';
import type { ChartEffectEdit } from './effects';
import type { ChartTextEdit } from './text';
import type { ChartThreeDEdit } from './three-d';

export function chartPointsToEmu(value: number): number {
	if (!Number.isFinite(value) || value < 0 || value * 12700 > Number.MAX_SAFE_INTEGER)
		throw new RangeError('Invalid chart size in points');
	return Math.round(value * 12700);
}

/** Hex plus transparency as an authored RGB colour. Invalid input is never silently clamped. */
export function chartDrawingColorFromHex(hex: string, transparency = 0): DrawingColor | null {
	const value = /^#?([0-9a-f]{6})$/iu.exec(hex)?.[1];
	if (!value || !Number.isFinite(transparency) || transparency < 0 || transparency > 100)
		return null;
	return {
		kind: 'srgb',
		value: value.toUpperCase(),
		transforms: transparency
			? [{ name: 'alpha', value: String(Math.round((100 - transparency) * 1000)) }]
			: [],
	};
}

function color(hex: string | undefined, transparency = 0): DrawingColor {
	const result = hex && chartDrawingColorFromHex(hex, transparency);
	if (!result) throw new RangeError('Invalid chart appearance colour');
	return result;
}

export function chartDrawingGradientFrom(value: ChartGradientFill): DrawingFill | null {
	if (value.stops.length < 2) return null;
	const stops: Extract<DrawingFill, { kind: 'gradient' }>['stops'] = [];
	for (const stop of value.stops) {
		const shade = chartDrawingColorFromHex(stop.color, 100 - (stop.opacity ?? 1) * 100);
		if (!shade || !Number.isFinite(stop.position) || stop.position < 0 || stop.position > 100)
			return null;
		stops.push({ position: stop.position, color: shade });
	}
	if (value.type === 'radial') {
		let focus = value.fillToRect;
		if (focus && !Object.values(focus).every(Number.isFinite)) return null;
		if (!focus && value.focalPoint) {
			const { x, y } = value.focalPoint;
			if (![x, y].every((channel) => Number.isFinite(channel) && channel >= 0 && channel <= 1))
				return null;
			focus = { l: x, t: y, r: 1 - x, b: 1 - y };
		}
		return {
			kind: 'gradient',
			stops,
			path: value.path ?? 'circle',
			...(focus ? { fillToRect: { ...focus } } : {}),
		};
	}
	if (!Number.isFinite(value.angle ?? 0) || (value.angle ?? 0) < 0 || (value.angle ?? 0) > 360)
		return null;
	return {
		kind: 'gradient',
		stops,
		angle: (value.angle ?? 0) % 360,
		scaled: value.scaled ?? false,
	};
}

export function chartFillFromAppearance(value: ChartFillAppearance): DrawingFill | undefined {
	switch (value.kind) {
		case 'automatic':
			return undefined;
		case 'none':
			return { kind: 'none' };
		case 'solid':
			return { kind: 'solid', color: color(value.color, value.transparency) };
		case 'gradient': {
			const fill = value.gradient && chartDrawingGradientFrom(value.gradient);
			if (!fill) throw new RangeError('Invalid chart gradient');
			return fill;
		}
		case 'pattern': {
			if (!value.pattern) throw new RangeError('A chart pattern needs its authored colours');
			return {
				kind: 'pattern',
				preset: value.pattern.preset,
				foreground: color(value.pattern.foreground),
				background: color(value.pattern.background),
			};
		}
		case 'picture': {
			const picture = value.picture;
			if (!picture || (!picture.relId && !picture.linkRelId))
				throw new RangeError('Picture data and relationships must be allocated by the host');
			const fill: Extract<DrawingFill, { kind: 'picture' }> = {
				kind: 'picture',
				...(picture.relId ? { relId: picture.relId } : {}),
				...(picture.linkRelId ? { linkRelId: picture.linkRelId } : {}),
			};
			if (picture.tile === undefined && picture.transparency === undefined) return fill;
			const root = parseXml(drawingFillXml(fill)!).documentElement;
			if (picture.tile) {
				const old = first(root, 'stretch', NS.a);
				if (old) root.removeChild(old);
				root.appendChild(root.ownerDocument.createElementNS(NS.a, 'a:tile'));
			}
			if (picture.transparency !== undefined) {
				if (
					!Number.isFinite(picture.transparency) ||
					picture.transparency < 0 ||
					picture.transparency > 100
				)
					throw new RangeError('Invalid picture transparency');
				const alpha = root.ownerDocument.createElementNS(NS.a, 'a:alphaModFix');
				alpha.setAttribute('amt', String(Math.round((100 - picture.transparency) * 1000)));
				first(root, 'blip', NS.a)!.appendChild(alpha);
			}
			return { ...fill, sourceXml: buildXml(root) };
		}
	}
}

export function chartLineFromAppearance(value: ChartLineAppearance): DrawingLine | undefined {
	if (value.kind === 'automatic' && Object.keys(value).length === 1) return undefined;
	const line: DrawingLine = {};
	if (value.kind === 'none') line.fill = { kind: 'none' };
	if (value.kind === 'solid')
		line.fill = { kind: 'solid', color: color(value.color, value.transparency) };
	if (value.kind === 'gradient') {
		const fill = value.gradient && chartDrawingGradientFrom(value.gradient);
		if (!fill) throw new RangeError('Invalid chart outline gradient');
		line.fill = fill;
	}
	if (value.width !== undefined) line.widthEmu = chartPointsToEmu(value.width);
	for (const key of ['dash', 'cap', 'compound'] as const)
		if (value[key] !== undefined) line[key] = value[key];
	if (value.join !== undefined) {
		if (!['round', 'bevel', 'miter'].includes(value.join))
			throw new RangeError('Invalid chart outline join');
		line.join = { kind: value.join as 'round' | 'bevel' | 'miter' };
	}
	if (value.beginArrow !== undefined) line.headEnd = { type: value.beginArrow };
	if (value.endArrow !== undefined) line.tailEnd = { type: value.endArrow };
	return line;
}

export function chartEffectEditsFromAppearance(value: ChartEffectsAppearance): ChartEffectEdit[] {
	const edits: ChartEffectEdit[] = [];
	if (value.shadow !== undefined) {
		const shadow = value.shadow;
		if (shadow === null) edits.push({ kind: 'outerShadow', remove: true });
		else {
			if (!Number.isFinite(shadow.angle)) throw new RangeError('Invalid shadow angle');
			edits.push({
				kind: 'outerShadow',
				color: color(shadow.color, shadow.transparency),
				blurRadiusEmu: chartPointsToEmu(shadow.blur),
				distanceEmu: chartPointsToEmu(shadow.distance),
				directionAngle: Math.round((((shadow.angle % 360) + 360) % 360) * 60000),
			});
		}
	}
	if (value.glow !== undefined)
		edits.push(
			value.glow === null
				? { kind: 'glow', remove: true }
				: {
						kind: 'glow',
						color: color(value.glow.color, value.glow.transparency),
						radiusEmu: chartPointsToEmu(value.glow.radius),
					},
		);
	if (value.softEdge !== undefined)
		edits.push(
			value.softEdge === null || value.softEdge === 0
				? { kind: 'softEdge', remove: true }
				: { kind: 'softEdge', radiusEmu: chartPointsToEmu(value.softEdge) },
		);
	return edits;
}

export function chartThreeDFromAppearance(value: ChartThreeDAppearance): ChartThreeDEdit {
	const edit: ChartThreeDEdit = {};
	for (const key of ['bevelTop', 'bevelBottom'] as const) {
		const bevel = value[key];
		if (bevel !== undefined)
			edit[key] =
				bevel === null
					? null
					: {
							preset: bevel.type,
							widthEmu: chartPointsToEmu(bevel.width),
							heightEmu: chartPointsToEmu(bevel.height),
						};
	}
	for (const [key, sizeKey, colorKey] of [
		['depth', 'depthEmu', 'depthColor'],
		['contour', 'contourWidthEmu', 'contourColor'],
	] as const) {
		const property = value[key];
		if (property === undefined) continue;
		edit[sizeKey] = property === null ? null : chartPointsToEmu(property.size);
		if (property === null) edit[colorKey] = null;
		else if (property.color !== undefined) edit[colorKey] = color(property.color);
	}
	if (value.material !== undefined) edit.material = value.material;
	if (value.lighting !== undefined) edit.lighting = value.lighting;
	return edit;
}

/** One text appearance may carry several independent effect edits. */
export function chartTextEditsFromAppearance(value: ChartTextAppearance): ChartTextEdit[] {
	const edit: ChartTextEdit = {};
	if (value.fill) edit.fill = chartFillFromAppearance(value.fill) ?? null;
	if (value.outline) edit.line = chartLineFromAppearance(value.outline) ?? null;
	for (const key of [
		'rotation',
		'direction',
		'anchor',
		'typeface',
		'eastAsiaTypeface',
		'sizePt',
		'bold',
		'italic',
	] as const)
		if (value[key] !== undefined) Object.assign(edit, { [key]: value[key] });
	return [
		edit,
		...chartEffectEditsFromAppearance(value.effects ?? {}).map((effect) => ({ effect })),
	];
}
