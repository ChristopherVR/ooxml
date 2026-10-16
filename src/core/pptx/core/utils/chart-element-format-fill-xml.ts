/**
 * Writes the fill choice of a chart element format (solid, gradient,
 * pattern, picture, none or automatic) into a `c:spPr`, an `a:ln` or a text
 * run's properties, in schema position and replacing every other member of
 * the choice group.
 *
 * Gradients go through the chart gradient builder (`buildChartGradFillXml`)
 * every other chart gradient writer uses.
 *
 * @module utils/chart-element-format-fill-xml
 */
import type {
	PptxChartElementFill,
	PptxChartElementLine,
	PptxChartPatternFill,
	PptxChartPictureFill,
	PptxChartTextFill,
	XmlObject,
} from '../types';
import { buildChartGradFillXml } from './chart-gradient-fill-xml';
import type { LocalName } from './chart-element-format-xml-util';
import {
	asObject,
	chartElementColorXml,
	keyOf,
	removeChildren,
	setOrdered,
	setPaintAlpha,
} from './chart-element-format-xml-util';
import { chartColorHex } from './chart-color-choice';

/** `EG_FillProperties`. */
export const FILL_CHOICE = ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'];
/** `EG_LineFillProperties`. */
export const LN_FILL_CHOICE = ['noFill', 'solidFill', 'gradFill', 'pattFill'];

/** `a:pattFill` content. */
export function chartPatternFillXml(pattern: PptxChartPatternFill): XmlObject {
	return {
		'@_prst': pattern.preset,
		'a:fgClr': { 'a:srgbClr': { '@_val': chartColorHex(pattern.foreground) } },
		'a:bgClr': { 'a:srgbClr': { '@_val': chartColorHex(pattern.background) } },
	};
}

/**
 * `a:blipFill` content for a picture with a relationship id, or `undefined`
 * when the picture has not been given one yet (the save path embeds it first).
 */
export function chartPictureFillXml(picture: PptxChartPictureFill): XmlObject | undefined {
	if (!picture.relId) {
		return undefined;
	}
	const blip: XmlObject = { '@_r:embed': picture.relId };
	if (picture.transparency !== undefined && picture.transparency > 0) {
		const amt = Math.round((100 - Math.min(100, picture.transparency)) * 1000);
		blip['a:alphaModFix'] = { '@_amt': String(amt) };
	}
	return {
		'@_rotWithShape': '1',
		'a:blip': blip,
		...(picture.tile
			? {
					'a:tile': {
						'@_tx': '0',
						'@_ty': '0',
						'@_sx': '100000',
						'@_sy': '100000',
						'@_flip': 'none',
						'@_algn': 'tl',
					},
				}
			: { 'a:stretch': { 'a:fillRect': {} } }),
	};
}

type Paint = PptxChartElementFill | PptxChartElementLine | PptxChartTextFill;

/**
 * Apply `paint` (fill, line fill or text fill) to `node`, whose fill-choice
 * children are `choices` and whose children follow `order`.
 */
export function applyChartPaint(
	node: XmlObject,
	paint: Paint,
	choices: readonly string[],
	order: readonly string[],
	localName: LocalName,
): void {
	const set = (name: string, value: XmlObject) => {
		removeChildren(node, choices, localName);
		setOrdered(node, name, value, order, localName);
	};
	switch (paint.kind) {
		case 'automatic':
			removeChildren(node, choices, localName);
			return;
		case 'none':
			set('noFill', {});
			return;
		case 'gradient':
			if (paint.gradient && paint.gradient.stops.length >= 2) {
				set('gradFill', buildChartGradFillXml(paint.gradient));
			}
			return;
		case 'pattern':
			if ('pattern' in paint && paint.pattern && choices.includes('pattFill')) {
				set('pattFill', chartPatternFillXml(paint.pattern));
			}
			return;
		case 'picture': {
			const blip =
				'picture' in paint && paint.picture && choices.includes('blipFill')
					? chartPictureFillXml(paint.picture)
					: undefined;
			if (blip) {
				set('blipFill', blip);
			}
			return;
		}
		default:
			break;
	}
	if (paint.color) {
		set('solidFill', chartElementColorXml(paint.color, paint.transparency));
		return;
	}
	const solidKey = keyOf(node, 'solidFill', localName);
	const solid = solidKey ? asObject(node[solidKey]) : undefined;
	if (solid && solidKey && paint.transparency !== undefined) {
		const next = { ...solid };
		setPaintAlpha(next, paint.transparency, localName);
		node[solidKey] = next;
	}
}
