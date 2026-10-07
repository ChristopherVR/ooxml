import { expect, it } from 'vitest';
import native from './excel-chart-styles.json';
import { readChartStyle } from './read-style';
import { resolveChartStyleDefinition } from './style-definition';
import { CHART_COLOR_STYLE_NS } from './color-style';

for (const sample of native.cases)
	it(`reads native style ${sample.style} and its measured font sizes`, () => {
		const xml = sample.parts['xl/charts/style1.xml'];
		const style = readChartStyle(xml)!;
		expect(style.id).toBe(sample.style);
		expect(style.sourceXml).toBe(xml);
		expect(Object.keys(style.entries).length).toBeGreaterThan(20);
		const resolved = resolveChartStyleDefinition(style, (color) => color.value)!;
		if (sample.titleFontSize === null) expect(resolved.title?.fontSize).toBeUndefined();
		else expect(resolved.title?.fontSize).toBe(sample.titleFontSize);
		expect(resolved.categoryAxis?.fontSize).toBe(sample.axisFontSize);
		expect(resolved.legend?.fontSize).toBe(sample.legendFontSize);
		expect(style.entries.title?.fontRef?.index).toBe(
			[208, 212, 215].includes(sample.style) ? 'major' : 'minor',
		);
	});

it('keeps transforms, arbitrary style entries and opaque effects without flattening', () => {
	const xml = `<cs:chartStyle xmlns:cs="${CHART_COLOR_STYLE_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" id="999"><cs:trendline><cs:fontRef idx="major"><a:schemeClr val="accent1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:spPr><a:effectLst><a:glow rad="12700"/></a:effectLst></cs:spPr></cs:trendline></cs:chartStyle>`;
	const style = readChartStyle(xml)!;
	expect(style.entries.trendline?.fontRef?.color?.transforms).toEqual([
		{ name: 'lumMod', value: '65000' },
		{ name: 'lumOff', value: '35000' },
	]);
	expect(style.entries.trendline?.sourceXml).toContain('glow rad="12700"');
	expect(style.sourceXml).toBe(xml);
});

it('prefers explicit text/line/fill colors and keeps zero-width/no-fill overrides', () => {
	const style = readChartStyle(
		`<cs:chartStyle xmlns:cs="${CHART_COLOR_STYLE_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><cs:title><cs:fontRef idx="minor"><a:schemeClr val="accent1"/></cs:fontRef><cs:lnRef idx="2"><a:schemeClr val="accent2"/></cs:lnRef><cs:fillRef idx="1"><a:schemeClr val="accent3"/></cs:fillRef><cs:defRPr sz="1800" b="0" i="1"><a:solidFill><a:srgbClr val="123456"/></a:solidFill></cs:defRPr><cs:spPr><a:noFill/><a:ln w="0"><a:noFill/></a:ln></cs:spPr></cs:title></cs:chartStyle>`,
	)!;
	expect(resolveChartStyleDefinition(style, (color) => color.value)?.title).toEqual({
		fontSize: 18,
		bold: false,
		italic: true,
		color: '123456',
		lineColor: 'none',
		fillColor: 'none',
	});
	expect(style.entries.title?.line?.widthEmu).toBe(0);
});

it('refuses a color-style part or the wrong namespace', () => {
	expect(readChartStyle(`<cs:colorStyle xmlns:cs="${CHART_COLOR_STYLE_NS}"/>`)).toBeUndefined();
	expect(readChartStyle('<chartStyle xmlns="urn:other"/>')).toBeUndefined();
});
