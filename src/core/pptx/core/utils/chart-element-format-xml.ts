/**
 * Writes a {@link PptxChartElementFormat} (the chart Format panes' Fill,
 * Border/Line and Effects sections) into a ChartML `c:spPr`.
 *
 * Each section present in the format replaces only its own part of the node
 * (the fill choice, the `a:ln` element, the modelled effects); everything
 * else in the authored `c:spPr` (`a:xfrm`, geometry, `a:scene3d`, `a:sp3d`,
 * extension lists, effects the panes do not model) passes through. Children
 * are kept in CT_ShapeProperties / CT_LineProperties / CT_EffectList order.
 *
 * @module utils/chart-element-format-xml
 */
import type {
	PptxChartElementEffects,
	PptxChartElementFill,
	PptxChartElementFormat,
	PptxChartElementLine,
	XmlObject,
} from '../types';
import { chartColorHex } from './chart-color-choice';

type LocalName = (key: string) => string;

const SP_PR_ORDER = [
	'xfrm',
	'custGeom',
	'prstGeom',
	'noFill',
	'solidFill',
	'gradFill',
	'blipFill',
	'pattFill',
	'grpFill',
	'ln',
	'effectLst',
	'effectDag',
	'scene3d',
	'sp3d',
	'extLst',
];
const FILL_CHOICE = ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'];
const LN_ORDER = [
	'noFill',
	'solidFill',
	'gradFill',
	'pattFill',
	'prstDash',
	'custDash',
	'round',
	'bevel',
	'miter',
	'headEnd',
	'tailEnd',
	'extLst',
];
const LN_FILL_CHOICE = ['noFill', 'solidFill', 'gradFill', 'pattFill'];
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

const EMU_PER_POINT = 12700;

function keyOf(node: XmlObject, name: string, localName: LocalName): string | undefined {
	return Object.keys(node).find((key) => !key.startsWith('@_') && localName(key) === name);
}

function asObject(value: unknown): XmlObject | undefined {
	return value && typeof value === 'object' && !Array.isArray(value)
		? (value as XmlObject)
		: undefined;
}

/** Remove the children named in `names`. */
function removeChildren(node: XmlObject, names: readonly string[], localName: LocalName): void {
	for (const key of Object.keys(node)) {
		if (!key.startsWith('@_') && names.includes(localName(key))) {
			delete node[key];
		}
	}
}

/** Set `a:<name>` at its schema position in `order`, replacing an existing child of that name. */
function setOrdered(
	node: XmlObject,
	name: string,
	value: XmlObject[string],
	order: readonly string[],
	localName: LocalName,
): void {
	const existing = keyOf(node, name, localName);
	if (existing) {
		node[existing] = value;
		return;
	}
	const rankOf = (key: string): number => {
		if (key.startsWith('@_')) {
			return -1;
		}
		const r = order.indexOf(localName(key));
		return r < 0 ? order.length : r;
	};
	// Stable sort: attributes first, then children in schema order (unknown ones last).
	const entries = [...Object.entries(node), [`a:${name}`, value] as [string, XmlObject[string]]]
		.map((entry, index) => ({ entry, index }))
		.sort((a, b) => rankOf(a.entry[0]) - rankOf(b.entry[0]) || a.index - b.index)
		.map(({ entry }) => entry);
	for (const key of Object.keys(node)) {
		delete node[key];
	}
	for (const [key, child] of entries) {
		node[key] = child;
	}
}

function alphaOf(transparency: number | undefined): XmlObject | undefined {
	if (transparency === undefined || transparency <= 0) {
		return undefined;
	}
	const clamped = Math.min(100, Math.max(0, transparency));
	return { '@_val': String(Math.round((100 - clamped) * 1000)) };
}

/** `a:srgbClr` with an optional `a:alpha`. */
export function chartElementColorXml(color: string, transparency?: number): XmlObject {
	const clr: XmlObject = { '@_val': chartColorHex(color) };
	const alpha = alphaOf(transparency);
	if (alpha) {
		clr['a:alpha'] = alpha;
	}
	return { 'a:srgbClr': clr };
}

/** Update only the `a:alpha` of whatever colour choice `paint` wraps. */
function setPaintAlpha(paint: XmlObject, transparency: number, localName: LocalName): void {
	const choiceKey = Object.keys(paint).find((key) => !key.startsWith('@_'));
	const choice = choiceKey ? asObject(paint[choiceKey]) : undefined;
	if (!choiceKey || !choice) {
		return;
	}
	const next: XmlObject = { ...choice };
	removeChildren(next, ['alpha'], localName);
	const alpha = alphaOf(transparency);
	if (alpha) {
		next['a:alpha'] = alpha;
	}
	paint[choiceKey] = next;
}

/** Apply `paint` (fill or line fill) to `node`, whose fill-choice children are `choices`. */
function applyPaint(
	node: XmlObject,
	paint: PptxChartElementFill | PptxChartElementLine,
	choices: readonly string[],
	order: readonly string[],
	localName: LocalName,
): void {
	if (paint.kind === 'automatic') {
		removeChildren(node, choices, localName);
		return;
	}
	if (paint.kind === 'none') {
		removeChildren(node, choices, localName);
		setOrdered(node, 'noFill', {}, order, localName);
		return;
	}
	if (paint.color) {
		removeChildren(node, choices, localName);
		setOrdered(
			node,
			'solidFill',
			chartElementColorXml(paint.color, paint.transparency),
			order,
			localName,
		);
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

function applyLine(spPr: XmlObject, line: PptxChartElementLine, localName: LocalName): void {
	const lnKey = keyOf(spPr, 'ln', localName);
	const ln: XmlObject = { ...(asObject(lnKey ? spPr[lnKey] : undefined) ?? {}) };
	applyPaint(ln, line, LN_FILL_CHOICE, LN_ORDER, localName);
	if (line.width !== undefined) {
		ln['@_w'] = String(Math.round(Math.max(0, line.width) * EMU_PER_POINT));
	}
	if (line.compound !== undefined) {
		ln['@_cmpd'] = line.compound;
	}
	if (line.cap !== undefined) {
		ln['@_cap'] = line.cap;
	}
	if (line.dash !== undefined) {
		removeChildren(ln, ['custDash'], localName);
		setOrdered(ln, 'prstDash', { '@_val': line.dash }, LN_ORDER, localName);
	}
	if (line.join !== undefined) {
		removeChildren(ln, ['round', 'bevel', 'miter'], localName);
		setOrdered(
			ln,
			line.join,
			line.join === 'miter' ? { '@_lim': '800000' } : {},
			LN_ORDER,
			localName,
		);
	}
	if (line.beginArrow !== undefined) {
		setOrdered(ln, 'headEnd', { '@_type': line.beginArrow }, LN_ORDER, localName);
	}
	if (line.endArrow !== undefined) {
		setOrdered(ln, 'tailEnd', { '@_type': line.endArrow }, LN_ORDER, localName);
	}
	if (Object.keys(ln).length === 0) {
		if (lnKey) {
			delete spPr[lnKey];
		}
		return;
	}
	setOrdered(spPr, 'ln', ln, SP_PR_ORDER, localName);
}

const pt = (value: number): string => String(Math.round(Math.max(0, value) * EMU_PER_POINT));

function applyEffects(
	spPr: XmlObject,
	effects: PptxChartElementEffects,
	localName: LocalName,
): void {
	const key = keyOf(spPr, 'effectLst', localName);
	const list: XmlObject = { ...(asObject(key ? spPr[key] : undefined) ?? {}) };
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
			delete spPr[key];
		}
		return;
	}
	setOrdered(spPr, 'effectLst', list, SP_PR_ORDER, localName);
}

/**
 * Return `existing` (a `c:spPr`, or `undefined` when the element has none)
 * with `format` applied, or `undefined` when nothing is left in it. With
 * `format.reset` the authored node is dropped first.
 */
export function applyChartElementFormatToSpPr(
	existing: XmlObject | undefined,
	format: PptxChartElementFormat,
	localName: LocalName,
): XmlObject | undefined {
	const spPr: XmlObject = format.reset ? {} : { ...(existing ?? {}) };
	if (format.fill) {
		applyPaint(spPr, format.fill, FILL_CHOICE, SP_PR_ORDER, localName);
	}
	if (format.line) {
		applyLine(spPr, format.line, localName);
	}
	if (format.effects) {
		applyEffects(spPr, format.effects, localName);
	}
	return Object.keys(spPr).length > 0 ? spPr : undefined;
}
