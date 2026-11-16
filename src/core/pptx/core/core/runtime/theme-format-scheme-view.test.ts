import { describe, expect, it } from 'vitest';

import { parseThemeFormatScheme } from '../../../../drawingml/format-scheme';
import { NS, first, parseXml } from '../../../../xml/index';
import type { ShapeStyle, XmlObject } from '../../types';
import { PptxHandlerRuntime } from './PptxHandlerRuntimeImplementation';
import { minimalThemeFormatScheme } from './PptxHandlerRuntimeSaveTheme';
import { drawingFormatSchemeOf } from './theme-format-scheme-view';

const A = NS.a;

/** A theme whose fill list mixes kinds out of the solid, gradient, pattern grouping. */
const THEME = `<a:theme xmlns:a="${A}" xmlns:r="${NS.r}" name="T"><a:themeElements>
<a:clrScheme name="C"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="111111"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
<a:accent1><a:srgbClr val="2468AC"/></a:accent1><a:accent2><a:srgbClr val="C03020"/></a:accent2>
<a:accent3><a:srgbClr val="00FF00"/></a:accent3><a:accent4><a:srgbClr val="0000FF"/></a:accent4>
<a:accent5><a:srgbClr val="FFFF00"/></a:accent5><a:accent6><a:srgbClr val="00FFFF"/></a:accent6>
<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>
<a:fontScheme name="F"><a:majorFont><a:latin typeface="A"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
<a:minorFont><a:latin typeface="B"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>
<a:fmtScheme name="M"><a:fillStyleLst>
<a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"/></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="50000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="0" scaled="0"/></a:gradFill>
<a:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>
<a:solidFill><a:schemeClr val="phClr"><a:alpha val="40000"/></a:schemeClr></a:solidFill>
</a:fillStyleLst><a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
<a:ln w="19050" cap="RND" cmpd="dbl"><a:solidFill><a:schemeClr val="accent2"/></a:solidFill><a:prstDash val="dash"/><a:round/></a:ln>
<a:ln w="28575"><a:noFill/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle>
<a:effectStyle><a:effectLst><a:outerShdw blurRad="38100" dist="19050" dir="5400000"><a:schemeClr val="phClr"><a:alpha val="60000"/></a:schemeClr></a:outerShdw></a:effectLst></a:effectStyle>
<a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
</a:fmtScheme></a:themeElements></a:theme>`;

class Runtime extends PptxHandlerRuntime {
	public load(xml: string) {
		const model = this.parseThemePartXml(xml);
		this.themeColorMap = { ...model.colorMap };
		this.themeFormatScheme = this.parseThemePartXml(xml).formatScheme;
		return this.themeFormatScheme!;
	}
	public ref(kind: 'fill' | 'line' | 'effect', node: XmlObject): ShapeStyle {
		const style: ShapeStyle = {};
		if (kind === 'fill') this.resolveThemeFillRef(node, style);
		else if (kind === 'line') this.resolveThemeLineRef(node, style);
		else this.resolveThemeEffectRef(node, style);
		return style;
	}
	public build(path: string): string | undefined {
		return this.buildThemeXml(path, undefined);
	}
}

const accent2 = (idx: number): XmlObject => ({
	'@_idx': String(idx),
	'a:schemeClr': { '@_val': 'accent2' },
});

describe('theme format scheme on the shared DrawingML model', () => {
	it('keeps the fill list in document order, a picture entry at its index', () => {
		const scheme = new Runtime().load(THEME);
		expect(scheme.name).toBe('M');
		expect(scheme.fillStyles.map((style) => style.kind)).toEqual(['gradient', 'picture', 'solid']);
		expect(scheme.fillStyles[2]).toMatchObject({ color: '#2468AC', opacity: 0.4 });
		expect(scheme.lineStyles[1]).toMatchObject({
			width: 2,
			color: '#C03020',
			dash: 'dash',
			lineJoin: 'round',
			lineCap: 'rnd',
			compoundLine: 'dbl',
		});
		expect(scheme.effectStyles[1]).toMatchObject({ shadowColor: '#2468AC', shadowOpacity: 0.6 });
	});

	it('substitutes the reference colour for phClr through the shared resolvers', () => {
		const runtime = new Runtime();
		runtime.load(THEME);
		expect(runtime.ref('fill', accent2(3))).toMatchObject({
			fillMode: 'solid',
			fillColor: '#C03020',
			fillOpacity: 0.4,
		});
		expect(runtime.ref('line', accent2(1))).toMatchObject({
			strokeColor: '#C03020',
			strokeWidth: 1,
		});
		expect(runtime.ref('effect', accent2(2))).toMatchObject({
			shadowColor: '#C03020',
			shadowOpacity: 0.6,
		});
		const gradient = runtime.ref('fill', accent2(1));
		expect(gradient.fillMode).toBe('gradient');
		expect(gradient.fillGradientStops?.[0]?.color).toBe('#C03020');
	});

	it('falls back to the reference colour for a picture entry', () => {
		const runtime = new Runtime();
		runtime.load(THEME);
		expect(runtime.ref('fill', accent2(2))).toMatchObject({
			fillMode: 'theme',
			fillColor: '#C03020',
		});
	});

	it('gives the neutral matrices back', () => {
		const runtime = new Runtime();
		const neutral = drawingFormatSchemeOf(runtime.load(THEME));
		const direct = parseThemeFormatScheme(
			first(first(parseXml(THEME).documentElement, 'themeElements', A), 'fmtScheme', A),
		)!;
		expect(neutral.fillStyles.map((style) => style.fill)).toEqual(
			direct.fillStyles.map((style) => style.fill),
		);
		expect(neutral.lineStyles).toEqual(direct.lineStyles);
		expect(neutral.effectStyles.map((style) => style.effects)).toEqual(
			direct.effectStyles.map((style) => style.effects),
		);
	});

	it('writes a schema-valid minimal format scheme for a part that has none', () => {
		const runtime = new Runtime();
		const xml = runtime.build('ppt/theme/theme9.xml')!;
		const fmt = first(first(parseXml(xml).documentElement, 'themeElements', A), 'fmtScheme', A);
		const scheme = parseThemeFormatScheme(fmt)!;
		expect(scheme.fillStyles).toHaveLength(3);
		expect(scheme.lineStyles.map((line) => line.widthEmu)).toEqual([6350, 12700, 19050]);
		expect(scheme.effectStyles).toHaveLength(3);
		expect(scheme.backgroundFillStyles).toHaveLength(3);
		expect(minimalThemeFormatScheme('X').name).toBe('X');
		expect(xml.match(/xmlns:a=/g)).toHaveLength(1);
	});
});
