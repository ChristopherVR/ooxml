import { expect, it } from 'vitest';
import { parseChartExSpace } from '../../../chart';
import { chartExFromNeutral } from './chart-ex-from-neutral';
it('adopts colorVal with sparse point identity, source cache and explicit fill', () => {
	const cache = '<cx:geoCache><cx:clear/></cx:geoCache>';
	const space = parseChartExSpace(
		`<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><cx:chartData><cx:data id="0"><cx:strDim type="cat"><cx:lvl ptCount="3"><cx:pt idx="2">Australia</cx:pt></cx:lvl></cx:strDim><cx:numDim type="colorVal"><cx:lvl ptCount="3"><cx:pt idx="2">42</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="regionMap"><cx:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></cx:spPr><cx:dataId val="0"/><cx:layoutPr><cx:geography>${cache}</cx:geography></cx:layoutPr></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>`,
	).chartExSpace;
	const result = chartExFromNeutral(space, {
		categories: [],
		series: [{ name: 'Value', values: [], color: '#4472C4' }],
	});
	expect(result.categories).toEqual(['Australia']);
	expect(result.series[0]!.values).toEqual([42]);
	expect(result.series[0]!.regionMapOptions).toMatchObject({
		fixedColor: '#4472C4',
		categorySourceIndices: [2],
		valueSourceIndices: [2],
	});
	expect(result.series[0]!.regionMapOptions?.geographyCacheXml).toContain('geoCache');
});
