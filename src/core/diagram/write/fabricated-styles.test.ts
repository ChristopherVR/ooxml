/**
 * The fabricated colour and quick style parts define every style label a built-in layout names,
 * so a background shape (the arrow behind Continuous Block Process) is not drawn without colours.
 */
import { describe, expect, it } from 'vitest';
import {
	buildFabricatedColorsXml,
	buildFabricatedQuickStyleXml,
	layoutStyleLabels,
} from './fabricated-styles';

const labelsOf = (xml: string) =>
	[...xml.matchAll(/<dgm:styleLbl name="([^"]+)">/gu)].map((m) => m[1]);
const label = (xml: string, name: string) =>
	new RegExp(`<dgm:styleLbl name="${name}">.*?</dgm:styleLbl>`, 'su').exec(xml)?.[0] ?? '';

describe('fabricated style parts', () => {
	it('reads the style labels a layout definition names, once each', () => {
		const layout =
			'<dgm:layoutDef><dgm:shape type="rightArrow" styleLbl="bgShp"/><dgm:layoutNode styleLbl="node1"/><dgm:layoutNode styleLbl="node1"/></dgm:layoutDef>';
		expect(layoutStyleLabels(layout)).toEqual(['bgShp', 'node1']);
	});

	it('keeps the standard labels when no layout adds any', () => {
		expect(labelsOf(buildFabricatedColorsXml('monochromatic2'))).toEqual([
			'node0',
			'node1',
			'parChTrans1D2',
			'sibTrans2D1',
		]);
	});

	it('colours a background label as a light tint of the scheme colour', () => {
		const colors = buildFabricatedColorsXml('monochromatic2', ['bgShp', 'node1']);
		expect(labelsOf(colors)).toEqual(['node0', 'node1', 'parChTrans1D2', 'sibTrans2D1', 'bgShp']);
		expect(label(colors, 'bgShp')).toContain(
			'<dgm:fillClrLst meth="repeat"><a:schemeClr val="accent2"><a:tint val="40000"/></a:schemeClr></dgm:fillClrLst>',
		);
		const style = buildFabricatedQuickStyleXml('flat', ['bgShp']);
		expect(label(style, 'bgShp')).toContain('<a:lnRef idx="0">');
	});
});
