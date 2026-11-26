import { EMU_PER_PIXEL } from '../units/constants';
import type { XmlElement } from '../xml/index';
import type { DrawingReflection } from './effect-model';
import { appendSvg } from './svg-elements';

const bounded = (value: number, min: number, max: number) =>
	Number.isFinite(value) ? Math.min(max, Math.max(min, value)) : min;
const fraction = (value: number) => bounded(value / 100000, 0, 1);

/** Paints the vertically flipped, fading reflection used by the Word theme style matrix. */
export function appendShapeReflection(
	defs: XmlElement,
	shape: XmlElement,
	id: string,
	width: number,
	height: number,
	effect: DrawingReflection,
): void {
	const parent = shape.parentElement;
	if (!parent) return;
	const scaleY = bounded(Math.abs(effect.scaleY ?? -100000) / 100000, 0, 2);
	const distance = bounded((effect.distanceEmu ?? 0) / EMU_PER_PIXEL, 0, 100);
	const blur = bounded((effect.blurRadiusEmu ?? 0) / EMU_PER_PIXEL / 2, 0, 25);
	const start = fraction(effect.startPosition ?? 0);
	const end = Math.max(start, fraction(effect.endPosition ?? 100000));
	const top = height + distance;
	const bottom = top + height * scaleY;
	if (top >= bottom) return;
	const pad = Math.max(2, Math.ceil(blur * 3));
	const region = { x: -pad, y: top - pad, width: width + pad * 2, height: bottom - top + pad * 2 };
	const fade = appendSvg(defs, 'linearGradient', {
		id: `${id}-fade`,
		gradientUnits: 'userSpaceOnUse',
		x1: 0,
		x2: 0,
		y1: top,
		y2: bottom,
	});
	appendSvg(fade, 'stop', { offset: start, 'stop-opacity': fraction(effect.startAlpha ?? 50000) });
	appendSvg(fade, 'stop', { offset: end, 'stop-opacity': fraction(effect.endAlpha ?? 0) });
	// Use black stops with an alpha mask so a light or dark shape fades the same way.
	const mask = appendSvg(defs, 'mask', {
		id: `${id}-mask`,
		maskUnits: 'userSpaceOnUse',
		style: 'mask-type:alpha',
		...region,
	});
	appendSvg(mask, 'rect', { ...region, fill: `url(#${id}-fade)` });
	const masked = appendSvg(parent, 'g', { mask: `url(#${id}-mask)` });
	parent.insertBefore(masked, shape);
	const flipped = appendSvg(masked, 'g', {
		transform: `translate(0 ${bottom}) scale(1 -${scaleY})`,
	});
	if (blur > 0) {
		const filter = appendSvg(defs, 'filter', {
			id: `${id}-blur`,
			x: '-50%',
			y: '-50%',
			width: '200%',
			height: '200%',
		});
		appendSvg(filter, 'feGaussianBlur', { stdDeviation: blur });
		flipped.setAttribute('filter', `url(#${id}-blur)`);
	}
	const drawn = shape.cloneNode(true);
	if (
		drawn.nodeType !== 1 ||
		!('removeAttribute' in drawn) ||
		typeof drawn.removeAttribute !== 'function'
	)
		return;
	drawn.removeAttribute('filter');
	flipped.appendChild(drawn);
}
