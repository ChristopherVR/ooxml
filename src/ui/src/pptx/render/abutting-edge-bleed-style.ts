/**
 * The style patch that grows a lower box under its abutting upper box (see
 * `abutting-edge-bleed.ts` for why). A binding merges it LAST over the
 * element's own container + fill style; it is empty unless the element has a
 * bleed and that style paints the fill as the box's own background.
 *
 * The box grows by {@link ABUTTING_EDGE_BLEED_PX} on each grown side. A plain
 * `background-color` simply fills the larger box. A `background-image` (a
 * gradient) keeps its authored size and position, and each grown side gets
 * one more layer under it: the fill's colour along that edge (or, for a fill
 * that runs along the edge, the fill image itself) reaching one pixel into
 * the authored box, so the boundary pixel is covered by a single primitive.
 */
import type { PptxElement } from 'ooxml-core/pptx';

import type { AbuttingEdgeBleed, AbuttingEdgeSide } from './abutting-edge-bleed';
import { ABUTTING_EDGE_BLEED_PX } from './abutting-edge-bleed';
import type { CssStyleMap } from './element-style-transform';

/** Keys that paint something other than the fill, or reposition it. */
const BLOCKING_KEYS = [
	'background',
	'backgroundSize',
	'backgroundPosition',
	'border',
	'borderWidth',
	'borderTop',
	'borderRight',
	'borderBottom',
	'borderLeft',
	'borderRadius',
	'boxShadow',
	'filter',
	'clipPath',
	'mask',
	'maskImage',
	'outline',
] as const;

const isSet = (value: unknown): boolean =>
	value !== undefined &&
	value !== null &&
	value !== '' &&
	value !== 'none' &&
	value !== 'transparent' &&
	value !== 0 &&
	value !== '0' &&
	value !== '0px';

const fmt = (n: number): string => `${Math.round(n * 1000) / 1000}px`;

const SIDES: readonly AbuttingEdgeSide[] = ['left', 'right', 'top', 'bottom'];

const kebab = (key: string): string => key.replace(/[A-Z]/gu, (c) => `-${c.toLowerCase()}`);

/**
 * The patch for `element` with `bleed`, given the style the binding has built
 * for its box (`painted`: camelCase or kebab-case keys, numbers or CSS
 * strings). The patch writes a background key in the spelling `painted` uses,
 * so it replaces the binding's own value instead of sitting beside it.
 */
export function abuttingEdgeBleedStyle(
	element: PptxElement,
	bleed: AbuttingEdgeBleed | undefined,
	painted: Readonly<Record<string, unknown>>,
): CssStyleMap {
	if (!bleed || SIDES.every((side) => bleed[side] === undefined)) return {};
	const read = (key: string): unknown => painted[key] ?? painted[kebab(key)];
	const kebabCase = Object.keys(painted).some((key) => key.startsWith('background-'));
	const spell = (key: string): string => (kebabCase ? kebab(key) : key);
	if (BLOCKING_KEYS.some((key) => isSet(read(key)))) return {};
	const repeat = read('backgroundRepeat');
	if (isSet(repeat) && repeat !== 'no-repeat') return {};
	// With no border (checked above) the border and padding boxes coincide, so
	// either origin or clip places the layers the same; any other value does not.
	for (const key of ['backgroundOrigin', 'backgroundClip']) {
		const value = read(key);
		if (isSet(value) && value !== 'border-box' && value !== 'padding-box') return {};
	}
	// Layers are placed from the border box: padding would shift a padding-box origin.
	const padded = ['padding', 'paddingTop', 'paddingRight', 'paddingBottom', 'paddingLeft'].some(
		(key) => isSet(read(key)),
	);
	if (padded && read('backgroundOrigin') !== 'border-box') return {};
	const image = read('backgroundImage');
	const hasImage = typeof image === 'string' && isSet(image);
	if (!hasImage && !isSet(read('backgroundColor'))) return {};

	const b = ABUTTING_EDGE_BLEED_PX;
	const grow = (side: AbuttingEdgeSide): number => (bleed[side] ? b : 0);
	const left = grow('left');
	const right = grow('right');
	const top = grow('top');
	const bottom = grow('bottom');
	const outerWidth = element.width + left + right;
	const outerHeight = element.height + top + bottom;
	const style: CssStyleMap = {
		left: fmt(element.x - left),
		top: fmt(element.y - top),
		width: fmt(outerWidth),
		height: fmt(outerHeight),
	};
	if (!hasImage) return style;

	const layers = [image];
	const sizes = [`${fmt(element.width)} ${fmt(element.height)}`];
	const positions = [`${fmt(left)} ${fmt(top)}`];
	for (const side of SIDES) {
		const paint = bleed[side];
		if (!paint) continue;
		const horizontal = side === 'left' || side === 'right';
		const across = b + 1;
		const x = side === 'right' ? left + element.width - 1 : 0;
		const y = side === 'bottom' ? top + element.height - 1 : 0;
		if (paint.color !== undefined) {
			layers.push(`linear-gradient(${paint.color}, ${paint.color})`);
			sizes.push(
				horizontal ? `${fmt(across)} ${fmt(outerHeight)}` : `${fmt(outerWidth)} ${fmt(across)}`,
			);
			positions.push(`${fmt(x)} ${fmt(y)}`);
		} else {
			// The fill runs along this edge: the same image, aligned with the
			// authored one along the edge, carries it into the strip.
			layers.push(image);
			sizes.push(
				horizontal
					? `${fmt(across)} ${fmt(element.height)}`
					: `${fmt(element.width)} ${fmt(across)}`,
			);
			positions.push(horizontal ? `${fmt(x)} ${fmt(top)}` : `${fmt(left)} ${fmt(y)}`);
		}
	}
	style[spell('backgroundImage')] = layers.join(', ');
	style[spell('backgroundSize')] = sizes.join(', ');
	style[spell('backgroundPosition')] = positions.join(', ');
	style[spell('backgroundRepeat')] = 'no-repeat';
	return style;
}
