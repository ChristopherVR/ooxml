import { expect, it } from 'vitest';
import { CHART_COLOR_PALETTES } from './color-palettes';
import { chartColorStyleId, chartColorStyleXml, CHART_COLOR_STYLE_NS } from './color-style';

for (const palette of CHART_COLOR_PALETTES)
	it(`writes native color-style ${palette.id} metadata and preserves extensions`, () => {
		const source = `<cs:colorStyle xmlns:cs="${CHART_COLOR_STYLE_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:v="urn:test" id="10" meth="cycle" custom="keep"><a:schemeClr val="accent1"/><cs:extLst><v:data value="keep"/></cs:extLst></cs:colorStyle>`;
		const xml = chartColorStyleXml(palette, source);
		expect(chartColorStyleId(xml)).toBe(palette.id);
		expect(xml).toContain('custom="keep"');
		expect(xml).toContain('value="keep"');
		expect(xml).toContain(`meth="${palette.meth}"`);
	});

it('does not infer a native palette from missing or invalid ids', () => {
	for (const id of ['', ' id=""', ' id="bad"'])
		expect(
			chartColorStyleId(`<cs:colorStyle xmlns:cs="${CHART_COLOR_STYLE_NS}"${id}/>`),
		).toBeUndefined();
});
