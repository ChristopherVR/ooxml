/**
 * Writes a {@link PptxChartElementFormat} (the chart Format panes' Fill,
 * Border/Line, Effects and 3-D Format sections) into a ChartML `c:spPr`.
 *
 * Each section present in the format replaces only its own part of the node
 * (the fill choice, the `a:ln` element, the modelled effects, `a:sp3d` and
 * the light rig); everything else in the authored `c:spPr` (`a:xfrm`,
 * geometry, the camera, extension lists, effects the panes do not model)
 * passes through. Children are kept in CT_ShapeProperties /
 * CT_LineProperties / CT_EffectList order. Text Options live in `c:txPr`
 * and are written by `chart-element-format-text-xml.ts`.
 *
 * @module utils/chart-element-format-xml
 */
import type { PptxChartElementFormat, PptxChartElementLine, XmlObject } from '../types';
import { applyChartThreeD } from './chart-element-format-3d-xml';
import { applyChartEffects } from './chart-element-format-effects-xml';
import { FILL_CHOICE, LN_FILL_CHOICE, applyChartPaint } from './chart-element-format-fill-xml';
import type { LocalName } from './chart-element-format-xml-util';
import { asObject, keyOf, pt, removeChildren, setOrdered } from './chart-element-format-xml-util';

export { chartElementColorXml } from './chart-element-format-xml-util';

const SP_PR_ORDER = [
	'xfrm',
	'custGeom',
	'prstGeom',
	...FILL_CHOICE,
	'ln',
	'effectLst',
	'effectDag',
	'scene3d',
	'sp3d',
	'extLst',
];
const LN_ORDER = [
	...LN_FILL_CHOICE,
	'prstDash',
	'custDash',
	'round',
	'bevel',
	'miter',
	'headEnd',
	'tailEnd',
	'extLst',
];

function applyLine(spPr: XmlObject, line: PptxChartElementLine, localName: LocalName): void {
	const lnKey = keyOf(spPr, 'ln', localName);
	const ln: XmlObject = { ...(asObject(lnKey ? spPr[lnKey] : undefined) ?? {}) };
	applyChartPaint(ln, line, LN_FILL_CHOICE, LN_ORDER, localName);
	if (line.width !== undefined) {
		ln['@_w'] = pt(line.width);
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
		applyChartPaint(spPr, format.fill, FILL_CHOICE, SP_PR_ORDER, localName);
	}
	if (format.line) {
		applyLine(spPr, format.line, localName);
	}
	if (format.effects) {
		applyChartEffects(spPr, format.effects, SP_PR_ORDER, localName);
	}
	if (format.threeD) {
		applyChartThreeD(spPr, format.threeD, localName, (name, value) => {
			if (value) {
				setOrdered(spPr, name, value, SP_PR_ORDER, localName);
			} else {
				removeChildren(spPr, [name], localName);
			}
		});
	}
	return Object.keys(spPr).length > 0 ? spPr : undefined;
}
