/**
 * Writes the Effects section of a chart element format (outer shadow, glow,
 * soft edges) into the `a:effectLst` of a `c:spPr` or of a text run's
 * properties, keeping effects the panes do not model.
 *
 * @module utils/chart-element-format-effects-xml
 */
import type { PptxChartElementEffects, XmlObject } from '../types';
import type { LocalName } from './chart-element-format-xml-util';
import {
	asObject,
	chartElementColorXml,
	keyOf,
	pt,
	removeChildren,
	setOrdered,
} from './chart-element-format-xml-util';

const EFFECT_ORDER = [
	'blur',
	'fillOverlay',
	'glow',
	'innerShdw',
	'outerShdw',
	'prstShdw',
	'reflection',
	'softEdge',
];

/** Apply `effects` to `parent/a:effectLst`, whose position follows `parentOrder`. */
export function applyChartEffects(
	parent: XmlObject,
	effects: PptxChartElementEffects,
	parentOrder: readonly string[],
	localName: LocalName,
): void {
	const key = keyOf(parent, 'effectLst', localName);
	const list: XmlObject = { ...(asObject(key ? parent[key] : undefined) ?? {}) };
	if (effects.shadow !== undefined) {
		removeChildren(list, ['outerShdw'], localName);
		if (effects.shadow) {
			const s = effects.shadow;
			setOrdered(
				list,
				'outerShdw',
				{
					'@_blurRad': pt(s.blur),
					'@_dist': pt(s.distance),
					'@_dir': String(Math.round((((s.angle % 360) + 360) % 360) * 60000)),
					'@_algn': 'tl',
					'@_rotWithShape': '0',
					...chartElementColorXml(s.color, s.transparency),
				},
				EFFECT_ORDER,
				localName,
			);
		}
	}
	if (effects.glow !== undefined) {
		removeChildren(list, ['glow'], localName);
		if (effects.glow) {
			setOrdered(
				list,
				'glow',
				{
					'@_rad': pt(effects.glow.radius),
					...chartElementColorXml(effects.glow.color, effects.glow.transparency),
				},
				EFFECT_ORDER,
				localName,
			);
		}
	}
	if (effects.softEdge !== undefined) {
		removeChildren(list, ['softEdge'], localName);
		if (effects.softEdge !== null && effects.softEdge > 0) {
			setOrdered(list, 'softEdge', { '@_rad': pt(effects.softEdge) }, EFFECT_ORDER, localName);
		}
	}
	if (Object.keys(list).length === 0) {
		if (key) {
			delete parent[key];
		}
		return;
	}
	setOrdered(parent, 'effectLst', list, parentOrder, localName);
}
