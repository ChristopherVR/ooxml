/**
 * Field changes of the newer chart Format pane sections (gradient details,
 * pattern and picture fills, 3-D Format and Text Options), turned into a
 * format patch for `withChartElementFormat`.
 *
 * @module render/chart-format/chart-format-pane-edit-extra
 */
import type {
	ChartElementFormatPatch,
	ChartElementFormatView,
	PptxChartBevel,
	PptxChartBevelType,
	PptxChartElementText,
	PptxChartGradientFill,
} from 'ooxml-core/pptx';

import type { ChartFormatFieldValue } from './chart-format-pane-types';

/**
 * PowerPoint applies a texture when Picture or texture fill is chosen with no
 * picture yet; this small tiled parchment stands in for its default.
 */
export const DEFAULT_CHART_TEXTURE =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAIAAACQkWg2AAABoUlEQVR42j3S504CQRQF4Hn/19BojCUxMYiiomIBFKmCIh1sgL389sMxJpPZnXPvKXd2QyW70qlt37UOn0f5m/Jmq5L8HJcm/dN2dev25mA6yObSc+PeSa++E6vhvp2xvPUbu3atjmpozdIGGpyWhQwPHhYHBzU7zvB6l+3wOo3sJRpqgwQHYi+3Z93LFJSqAj0ZRIJoeL07Z/7n8DTMqT12j9vVzZjn/eFi1Nx/6BzpsKMRZt5vpIgG7F59Ow4QTY2I/z2toL3dF3DKp8t0NTAMylRRo7DEuuPccub35+GcaX08FqkEA+izWHGnhKAsAO14y/bB1R5cTyDjTI+7s0itSgIBrhUHAneN5IGBarOUiMMVMoscBCVGwmXIgyAnOXPW8msB9WtSbhTW+ZgEEyK0DCS18jGSLxPB0Kklncmz06Gbkr7I91I8XlICCsJ5NrTEatjcoZpI0KvmViVUjbdnEllm1wr1ackIKrpUo+aemgHscBeKz9wK8REHRTNWHFckO3L8ZZjQnf0av//P7B7PDxf+/2o1IA4HNEeiGij+AIB5gfnxhkAIAAAAAElFTkSuQmCC';

/** A light-to-colour gradient for a line or text switched to Gradient. */
export function defaultChartGradient(color: string | undefined): PptxChartGradientFill {
	const base = color && /^#?[0-9a-f]{6}$/iu.test(color) ? color.replace(/^#?/u, '#') : '#4472C4';
	return {
		type: 'linear',
		angle: 0,
		stops: [
			{ color: '#FFFFFF', position: 0 },
			{ color: base.toUpperCase(), position: 100 },
		],
	};
}

const clamp = (n: number, min: number, max: number) => Math.min(max, Math.max(min, n));

function num(value: ChartFormatFieldValue): number | null {
	if (value === '' || typeof value === 'boolean') return null;
	const n = typeof value === 'number' ? value : Number(value);
	return Number.isFinite(n) ? n : null;
}

/** `gradient` with one detail field (`gradientType|Start|End|Angle`) changed, or `null`. */
export function editGradient(
	gradient: PptxChartGradientFill,
	detail: string,
	value: ChartFormatFieldValue,
): PptxChartGradientFill | null {
	const stops = gradient.stops.map((stop) => ({ ...stop }));
	switch (detail) {
		case 'gradientType':
			return value === 'linear' || value === 'radial'
				? value === 'linear'
					? { type: 'linear', stops, angle: gradient.angle ?? 0 }
					: { type: 'radial', stops }
				: null;
		case 'gradientStart':
			if (stops[0]) stops[0].color = String(value);
			return { ...gradient, stops };
		case 'gradientEnd':
			if (stops.length > 0) stops[stops.length - 1]!.color = String(value);
			return { ...gradient, stops };
		case 'gradientAngle': {
			const n = num(value);
			return n === null ? null : { ...gradient, angle: clamp(n, 0, 359) };
		}
		default:
			return null;
	}
}

/** Picture and pattern details of the Fill section. */
export function fillDetailPatch(
	view: ChartElementFormatView,
	field: string,
	value: ChartFormatFieldValue,
): ChartElementFormatPatch | null {
	const n = num(value);
	switch (field) {
		case 'fill.picture':
			return typeof value === 'string' && value.startsWith('data:image/')
				? { fill: { kind: 'picture', picture: { src: value } } }
				: null;
		case 'fill.pictureTransparency':
			return n === null
				? null
				: { fill: { kind: 'picture', picture: { transparency: clamp(n, 0, 100) } } };
		case 'fill.pictureTile':
			return { fill: { kind: 'picture', picture: { tile: Boolean(value) } } };
		case 'fill.patternPreset':
			return { fill: { kind: 'pattern', pattern: { preset: String(value) } } };
		case 'fill.patternForeground':
			return { fill: { kind: 'pattern', pattern: { foreground: String(value) } } };
		case 'fill.patternBackground':
			return { fill: { kind: 'pattern', pattern: { background: String(value) } } };
		default: {
			const [section, detail] = field.split('.');
			if ((section === 'fill' || section === 'line') && detail) {
				const current = section === 'fill' ? view.fill.gradient : view.line.gradient;
				const gradient = current && editGradient(current, detail, value);
				if (!gradient) return null;
				return section === 'fill'
					? { fill: { kind: 'gradient', gradient } }
					: { line: { kind: 'gradient', gradient } };
			}
			return null;
		}
	}
}

/** One 3-D Format field (`threeD.*`). */
export function threeDPatch(
	view: ChartElementFormatView,
	field: string,
	value: ChartFormatFieldValue,
): ChartElementFormatPatch | null {
	const n = num(value);
	const key = field.slice('threeD.'.length);
	const bevel = (which: 'bevelTop' | 'bevelBottom', prop: 'type' | 'width' | 'height') => {
		const current: PptxChartBevel = view.threeD[which] ?? { type: 'circle', width: 6, height: 6 };
		if (prop === 'type') {
			return value === 'none'
				? { threeD: { [which]: null } }
				: { threeD: { [which]: { ...current, type: String(value) as PptxChartBevelType } } };
		}
		return n === null ? null : { threeD: { [which]: { ...current, [prop]: clamp(n, 0, 1584) } } };
	};
	switch (key) {
		case 'bevelTop':
			return bevel('bevelTop', 'type');
		case 'bevelTopWidth':
			return bevel('bevelTop', 'width');
		case 'bevelTopHeight':
			return bevel('bevelTop', 'height');
		case 'bevelBottom':
			return bevel('bevelBottom', 'type');
		case 'bevelBottomWidth':
			return bevel('bevelBottom', 'width');
		case 'bevelBottomHeight':
			return bevel('bevelBottom', 'height');
		case 'depthSize':
		case 'contourSize': {
			if (n === null) return null;
			const which = key === 'depthSize' ? 'depth' : 'contour';
			const color = view.threeD[which]?.color;
			return {
				threeD: {
					[which]: n > 0 ? { size: clamp(n, 0, 1584), ...(color ? { color } : {}) } : null,
				},
			};
		}
		case 'depthColor':
		case 'contourColor': {
			const which = key === 'depthColor' ? 'depth' : 'contour';
			const size = view.threeD[which]?.size ?? (which === 'depth' ? 0 : 0);
			// A colour alone has nothing to paint until the size is set; keep it with a 1pt size.
			return { threeD: { [which]: { size: size > 0 ? size : 1, color: String(value) } } };
		}
		case 'material':
			return { threeD: { material: String(value) } };
		case 'lighting':
			return { threeD: { lighting: String(value) } };
		default:
			return null;
	}
}

type TextPatch = NonNullable<ChartElementFormatPatch['text']>;

/** One Text Options field (`text.*`). */
export function textPatch(
	view: ChartElementFormatView,
	field: string,
	value: ChartFormatFieldValue,
	defaults: {
		shadow: NonNullable<PptxChartElementText['effects']>['shadow'];
		glow: NonNullable<PptxChartElementText['effects']>['glow'];
	},
): ChartElementFormatPatch | null {
	const n = num(value);
	const text = String(value);
	const wrap = (patch: TextPatch): ChartElementFormatPatch => ({ text: patch });
	switch (field) {
		case 'text.fill.kind':
			if (text === 'gradient') {
				return wrap({
					fill: {
						kind: 'gradient',
						gradient: view.text.fill?.gradient ?? defaultChartGradient(view.text.fill?.color),
					},
				});
			}
			return text === 'none' || text === 'solid' || text === 'automatic'
				? wrap({
						fill: {
							kind: text,
							...(text === 'solid' && !view.text.fill?.color ? { color: '#000000' } : {}),
						},
					})
				: null;
		case 'text.fill.color':
			return wrap({ fill: { kind: 'solid', color: text } });
		case 'text.fill.transparency':
			return n === null ? null : wrap({ fill: { transparency: clamp(n, 0, 100) } });
		case 'text.outline.kind':
			return text === 'none' || text === 'solid' || text === 'automatic'
				? wrap({
						outline: {
							kind: text,
							...(text === 'solid' && !view.text.outline?.color ? { color: '#000000' } : {}),
						},
					})
				: null;
		case 'text.outline.color':
			return wrap({ outline: { kind: 'solid', color: text } });
		case 'text.outline.transparency':
			return n === null ? null : wrap({ outline: { transparency: clamp(n, 0, 100) } });
		case 'text.outline.width':
			return n === null ? null : wrap({ outline: { width: clamp(n, 0, 1584) } });
		case 'text.anchor':
			return text === 't' || text === 'ctr' || text === 'b' ? wrap({ anchor: text }) : null;
		case 'text.direction':
			return text === 'horz' || text === 'vert' || text === 'vert270' || text === 'wordArtVert'
				? wrap({ direction: text })
				: null;
		case 'text.rotation':
			return value === ''
				? wrap({ rotation: null })
				: n === null
					? null
					: wrap({ rotation: clamp(Math.round(n), -90, 90) });
		case 'text.shadow.enabled':
			return wrap({ effects: { shadow: value ? defaults.shadow : null } });
		case 'text.glow.enabled':
			return wrap({ effects: { glow: value ? defaults.glow : null } });
		default:
			break;
	}
	if (field.startsWith('text.fill.gradient') && view.text.fill?.gradient) {
		const gradient = editGradient(view.text.fill.gradient, field.slice('text.fill.'.length), value);
		return gradient ? wrap({ fill: { kind: 'gradient', gradient } }) : null;
	}
	const [, effect, prop] = field.split('.');
	if (effect === 'shadow' && view.text.effects?.shadow && prop) {
		const shadow = { ...view.text.effects.shadow };
		if (prop === 'color') shadow.color = text;
		else if (n === null) return null;
		else if (
			prop === 'transparency' ||
			prop === 'blur' ||
			prop === 'distance' ||
			prop === 'angle'
		) {
			shadow[prop] =
				prop === 'angle' ? clamp(n, 0, 359) : clamp(n, 0, prop === 'transparency' ? 100 : 200);
		}
		return wrap({ effects: { shadow } });
	}
	if (effect === 'glow' && view.text.effects?.glow && prop) {
		const glow = { ...view.text.effects.glow };
		if (prop === 'color') glow.color = text;
		else if (n === null) return null;
		else if (prop === 'radius' || prop === 'transparency') {
			glow[prop] = clamp(n, 0, prop === 'transparency' ? 100 : 150);
		}
		return wrap({ effects: { glow } });
	}
	return null;
}
