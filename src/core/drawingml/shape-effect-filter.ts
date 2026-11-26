import { EMU_PER_PIXEL } from '../units/constants';
import type { XmlElement } from '../xml/index';
import { resolveDrawingColor, type DrawingColorTheme } from './drawing-color';
import { resolveDrawingOuterShadow, type DrawingSvgShadow } from './drawing-shadow';
import type { DrawingEffectList } from './effect-model';
import { appendSvg } from './svg-elements';

const round = (value: number) => Math.round(value * 100) / 100;
const bound = (value: number) => round(Math.min(100, Math.max(-100, value)));

function shadowAttrs(shadow: DrawingSvgShadow) {
	return {
		dx: bound(shadow.dx),
		dy: bound(shadow.dy),
		stdDeviation: bound(shadow.blur / 2),
		'flood-color': shadow.color,
		'flood-opacity': round(shadow.opacity),
	};
}

/** Paints the ordinary shadow, glow and soft-edge effects supported by the shape renderer. */
export function appendShapeEffectFilter(
	defs: XmlElement,
	id: string,
	effects: DrawingEffectList | undefined,
	theme: DrawingColorTheme,
): boolean {
	let filter: XmlElement | undefined;
	const element = () =>
		(filter ??= appendSvg(defs, 'filter', {
			id,
			x: '-50%',
			y: '-50%',
			width: '200%',
			height: '200%',
		}));
	let previous = 'SourceGraphic';
	for (const [index, effect] of (effects?.effects ?? []).slice(0, 16).entries()) {
		const result = `effect${index}`;
		if (effect.kind === 'outerShadow' || effect.kind === 'innerShadow') {
			const shadow = resolveDrawingOuterShadow({ ...effect, kind: 'outerShadow' }, theme);
			if (!shadow) continue;
			const attrs = shadowAttrs(shadow);
			if (effect.kind === 'outerShadow') {
				appendSvg(element(), 'feDropShadow', { in: previous, ...attrs, result });
			} else {
				const node = element();
				appendSvg(node, 'feOffset', { in: 'SourceAlpha', dx: attrs.dx, dy: attrs.dy });
				appendSvg(node, 'feGaussianBlur', {
					stdDeviation: attrs.stdDeviation,
					result: 'shiftedAlpha',
				});
				appendSvg(node, 'feComposite', {
					in: 'SourceAlpha',
					in2: 'shiftedAlpha',
					operator: 'out',
					result: 'innerMask',
				});
				appendSvg(node, 'feFlood', {
					'flood-color': attrs['flood-color'],
					'flood-opacity': attrs['flood-opacity'],
				});
				appendSvg(node, 'feComposite', { in2: 'innerMask', operator: 'in' });
				appendSvg(node, 'feComposite', { in2: previous, operator: 'atop', result });
			}
		} else if (effect.kind === 'glow') {
			const color = effect.color && resolveDrawingColor(effect.color, theme);
			const radius = (effect.radiusEmu ?? 0) / EMU_PER_PIXEL;
			if (!color || !Number.isFinite(radius) || radius <= 0) continue;
			appendSvg(element(), 'feDropShadow', {
				in: previous,
				dx: 0,
				dy: 0,
				stdDeviation: bound(radius / 2),
				'flood-color': color.hex,
				'flood-opacity': round(color.alpha),
				result,
			});
		} else if (effect.kind === 'softEdge') {
			const radius = (effect.radiusEmu ?? 0) / EMU_PER_PIXEL;
			if (!Number.isFinite(radius) || radius <= 0) continue;
			const node = element();
			appendSvg(node, 'feMorphology', {
				in: 'SourceAlpha',
				operator: 'erode',
				radius: bound(radius / 2),
			});
			appendSvg(node, 'feGaussianBlur', { stdDeviation: bound(radius / 2), result: 'softAlpha' });
			appendSvg(node, 'feComposite', { in: previous, in2: 'softAlpha', operator: 'in', result });
		} else continue;
		previous = result;
	}
	return Boolean(filter);
}
