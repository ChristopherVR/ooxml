import { describe, expect, it } from 'vitest';

import { withChartElementFormat } from '../builders/sdk/chart-element-format-operations';
import type { PptxChartData, XmlObject } from '../types';
import { applyChartTextFormat, readChartTextFormat } from './chart-element-format-text-xml';

const local = (key: string) => key.replace(/^.*:/u, '');
const read = (owner: XmlObject) =>
	readChartTextFormat(
		owner,
		local,
		(node) => (node['a:solidFill'] ? { kind: 'solid', color: '#C00000' } : undefined),
		() => undefined,
	);

describe('chart Text Options XML', () => {
	it("treats PowerPoint's rot=-60000000 as automatic and reads a real angle", () => {
		expect(read({ 'c:txPr': { 'a:bodyPr': { '@_rot': '-60000000' } } })?.rotation).toBeUndefined();
		expect(read({ 'c:txPr': { 'a:bodyPr': { '@_rot': '-2700000', '@_vert': 'horz' } } })).toEqual({
			rotation: -45,
			direction: 'horz',
		});
	});

	it('creates c:txPr before the following children and writes every paragraph default', () => {
		const owner: XmlObject = { 'c:spPr': {}, 'c:crossAx': { '@_val': '1' } };
		applyChartTextFormat(
			owner,
			{ fill: { kind: 'solid', color: '#C00000' }, rotation: 30 },
			['crossAx'],
			local,
		);
		expect(Object.keys(owner)).toEqual(['c:spPr', 'c:txPr', 'c:crossAx']);
		const txPr = owner['c:txPr'] as XmlObject;
		expect(txPr['a:bodyPr']).toEqual({ '@_rot': '1800000' });
		expect(read(owner)).toMatchObject({ rotation: 30, fill: { kind: 'solid' } });
	});

	it('mirrors a solid text fill onto the flat text colours', () => {
		const data: PptxChartData = {
			chartType: 'bar',
			categories: ['A'],
			series: [{ name: 'S', values: [1] }],
			titleRuns: [{ text: 'Sales' }],
			axes: [{ axisType: 'valAx', axisId: 1 }],
		};
		const title = withChartElementFormat(
			data,
			{ kind: 'title' },
			{ text: { fill: { kind: 'solid', color: '#2F5597' } } },
		);
		expect(title.style?.titleFontColor).toBe('#2F5597');
		expect(title.titleRuns?.[0]?.color).toBe('#2F5597');
		const axis = withChartElementFormat(
			data,
			{ kind: 'axis', axisIndex: 0 },
			{ text: { fill: { kind: 'solid', color: '#C00000' } } },
		);
		expect(axis.axes?.[0]?.fontColor).toBe('#C00000');
		const cleared = withChartElementFormat(
			axis,
			{ kind: 'axis', axisIndex: 0 },
			{ text: { fill: { kind: 'automatic' } } },
		);
		expect(cleared.axes?.[0]?.fontColor).toBeUndefined();
	});
});
