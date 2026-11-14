/**
 * Run a measurement with every CSS transform inside a stage switched off, so
 * `getBoundingClientRect` and `Range.getClientRects` report the untransformed layout. Each
 * element's own transform is read first and handed over as a matrix about its origin, in
 * untransformed stage CSS px: painting reapplies it. The stage's own transform and its
 * ancestors' (zoom) are left alone: they only scale and move the stage, which `measure` undoes.
 * The live styles are put back before this returns; the swap is synchronous, so no frame is
 * painted in between.
 */
import { cssMatrix, cssOrigin } from './pdf-vector-css';
import type { Matrix } from './pdf-vector-display-list';
import { multiply } from './pdf-vector-display-list';

/** `local` is unset for a transform that is not flat (perspective, true 3-D). */
export type ElementTransforms = ReadonlyMap<Element, { local?: Matrix }>;

const PROPERTIES = ['transform', 'translate', 'rotate', 'scale'] as const;

/** The individual `translate`, `rotate` and `scale` properties as one matrix. */
function individualMatrix(style: CSSStyleDeclaration): Matrix | undefined {
	let m: Matrix = [1, 0, 0, 1, 0, 0];
	const translate = style.getPropertyValue('translate').trim();
	if (translate && translate !== 'none') {
		const [x = '0', y = '0', z] = translate.split(/\s+/);
		if (z || !/px$|^0$/.test(x) || !/px$|^0$/.test(y)) return undefined;
		m = multiply(m, [1, 0, 0, 1, Number.parseFloat(x), Number.parseFloat(y)]);
	}
	const rotate = style.getPropertyValue('rotate').trim();
	if (rotate && rotate !== 'none') {
		if (!/^-?[\d.]+deg$/.test(rotate)) return undefined;
		const a = (Number.parseFloat(rotate) * Math.PI) / 180;
		m = multiply(m, [Math.cos(a), Math.sin(a), -Math.sin(a), Math.cos(a), 0, 0]);
	}
	const scale = style.getPropertyValue('scale').trim();
	if (scale && scale !== 'none') {
		const [x = '1', y = x, z] = scale.split(/\s+/);
		if (z && z !== '1') return undefined;
		m = multiply(m, [Number.parseFloat(x), 0, 0, Number.parseFloat(y), 0, 0]);
	}
	return m;
}

const isIdentity = (m: Matrix) =>
	m[0] === 1 && m[1] === 0 && m[2] === 0 && m[3] === 1 && m[4] === 0 && m[5] === 0;

interface Saved {
	el: HTMLElement | SVGElement;
	property: string;
	value: string;
	priority: string;
}

/** Viewport px to untransformed stage CSS px. */
export interface StageMeasure {
	originX: number;
	originY: number;
	/** Viewport px per stage CSS px (the stage's and its ancestors' scale). */
	scale: number;
}

/** Call `read` with the transforms inside `stage` off (see the module comment). */
export function neutraliseTransforms<T>(
	stage: HTMLElement,
	view: Window,
	read: (transforms: ElementTransforms, measure: StageMeasure) => T,
): T {
	const found: { el: Element; local?: Matrix; origin: string }[] = [];
	const boxes = Array.from(stage.querySelectorAll('*')).filter(
		// SVG content keeps its own (attribute) transforms: the SVG converter applies them.
		(el) => !el.parentElement?.closest('svg'),
	);
	for (const el of boxes) {
		const style = view.getComputedStyle(el);
		const main = cssMatrix(style.transform);
		const individual = individualMatrix(style);
		if (main && individual && isIdentity(main) && isIdentity(individual)) continue;
		const local = main && individual ? multiply(individual, main) : undefined;
		found.push({ el, ...(local ? { local } : {}), origin: style.transformOrigin });
	}
	const saved: Saved[] = [];
	for (const { el } of found) {
		const styled = el as HTMLElement;
		if (!styled.style) continue;
		for (const property of PROPERTIES) {
			saved.push({
				el: styled,
				property,
				value: styled.style.getPropertyValue(property),
				priority: styled.style.getPropertyPriority(property),
			});
			styled.style.setProperty(property, 'none', 'important');
		}
	}
	try {
		const stageRect = stage.getBoundingClientRect();
		const scale = stage.offsetWidth > 0 ? stageRect.width / stage.offsetWidth : 1;
		const measure: StageMeasure = {
			originX: stageRect.left,
			originY: stageRect.top,
			scale: scale > 0 ? scale : 1,
		};
		const transforms = new Map<Element, { local?: Matrix }>();
		for (const entry of found) {
			if (!entry.local) {
				transforms.set(entry.el, {});
				continue;
			}
			const rect = entry.el.getBoundingClientRect();
			const [ox0, oy0] = cssOrigin(
				entry.origin,
				rect.width / measure.scale,
				rect.height / measure.scale,
			);
			const ox = (rect.left - measure.originX) / measure.scale + ox0;
			const oy = (rect.top - measure.originY) / measure.scale + oy0;
			transforms.set(entry.el, {
				local: multiply(multiply([1, 0, 0, 1, ox, oy], entry.local), [1, 0, 0, 1, -ox, -oy]),
			});
		}
		return read(transforms, measure);
	} finally {
		for (const { el, property, value, priority } of saved.reverse()) {
			if (value) el.style.setProperty(property, value, priority);
			else el.style.removeProperty(property);
		}
	}
}
