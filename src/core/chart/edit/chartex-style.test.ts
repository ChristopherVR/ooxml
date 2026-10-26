import { expect, it } from 'vitest';
import { createChartExSpace } from '../ex/create';
import { chartExDataChanged } from '../ex/workbook-writes';
import { writeChartExSpace } from '../ex/write';
import { parseChartExSpace } from '../ex/parse';
import { applyChartExEditIntent } from './chartex-intent';
import { editChartExAxis } from './chartex-axis';
import { editChartExPointLabel } from './chartex-point-label';
import { editChartExStyle } from './chartex-style';
import { editChartSpace } from './draft';
import { editChartTextBody } from './text';
import { chartTextBodyFromXml } from '../write-shape';
const source = () =>
	createChartExSpace(
		{ categories: ['A', 'B'], series: [{ name: 'Sales', values: [1, 2] }] },
		'waterfall',
		{ title: 'Title', legend: { visible: true } },
	);
it('uses shared pattern, line, effect and text operations without changing chart data or family metadata', () => {
	const space = source();
	const pattern = applyChartExEditIntent(space, {
		op: 'pattern',
		target: { kind: 'series', seriesIndex: 0 },
		edit: { preset: 'diagCross' },
	});
	const line = applyChartExEditIntent(pattern, {
		op: 'lineOptions',
		target: { kind: 'axis', axisIndex: 1 },
		edit: { widthEmu: 25400, dash: 'dash' },
	});
	const effect = applyChartExEditIntent(line, {
		op: 'effect',
		target: { kind: 'legend' },
		edit: { kind: 'softEdge', radiusEmu: 12700 },
	});
	const text = applyChartExEditIntent(effect, {
		op: 'text',
		target: { kind: 'title' },
		edit: { sizePt: 24, typeface: 'Aptos', bold: true },
	});
	expect(chartExDataChanged(space, text)).toBe(false);
	const result = parseChartExSpace(writeChartExSpace(text));
	expect(result.issues).toEqual([]);
	expect(result.chartExSpace.plotArea.series[0]!.spPr!.fill).toMatchObject({
		kind: 'pattern',
		preset: 'diagCross',
	});
	expect(result.chartExSpace.plotArea.axes[1]!.spPr!.line).toMatchObject({
		widthEmu: 25400,
		dash: 'dash',
	});
	expect(result.chartExSpace.legend!.spPr!.effectsXml).toContain('softEdge');
	expect(result.chartExSpace.title!.tx!.rich!.sourceXml).toContain('sz="2400"');
});
it('patches axis semantics and point defaults/hidden state without replacing siblings', () => {
	const space = source();
	const axis = editChartExAxis(space, 1, {
		min: 0,
		max: 20,
		majorUnit: 5,
		title: 'Value axis',
		tickLabels: false,
	});
	const label = editChartExPointLabel(axis, 0, 0, {
		visibility: { value: true },
		font: { sizePt: 12 },
		deleted: true,
	});
	const second = editChartExPointLabel(label, 0, 1, { visibility: { categoryName: true } });
	expect(second.plotArea.series[0]!.dataLabels!.hidden).toEqual([0]);
	expect(second.plotArea.series[0]!.dataLabels!.labels).toHaveLength(2);
	expect(second.plotArea.axes[1]!.title!.tx!.rich!.text).toBe('Value axis');
	expect(second.plotArea.axes[1]!.tickLabels).toBeUndefined();
	expect(() => editChartExAxis(axis, 1, { max: -1 })).toThrow('maximum');
	expect(() => editChartExAxis(axis, 1, { majorUnit: 0 })).toThrow();
	expect(() => editChartExAxis(axis, 0, { min: 0 })).toThrow('value axis');
});
it('projects and copies rich axis title fonts while retaining authored text and theme metadata', () => {
	const space = source();
	const rich = chartTextBodyFromXml(
		'<a:bodyPr/><a:p><a:r><a:rPr sz="1200" b="0"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:latin typeface="Producer"/></a:rPr><a:t>Authored axis</a:t></a:r></a:p>',
	);
	space.plotArea.axes[1]!.title = { tx: { rich } };
	const edited = editChartExStyle(space, (snapshot) =>
		editChartSpace(snapshot, (draft) => {
			const title = draft.plotArea.axes[1]!.title!;
			title.tx!.rich = editChartTextBody(title.tx!.rich, { sizePt: 24, bold: true });
			title.txPr = editChartTextBody(title.txPr, { sizePt: 24, bold: true });
		}),
	);
	const title = edited.plotArea.axes[1]!.title!;
	expect(title.tx!.rich!.text).toBe('Authored axis');
	expect(title.tx!.rich!.sourceXml).toContain('sz="2400"');
	expect(title.tx!.rich!.sourceXml).toContain('b="1"');
	expect(title.tx!.rich!.sourceXml).toContain('typeface="Producer"');
	expect(title.tx!.rich!.sourceXml).toContain('schemeClr val="accent1"');
	expect(title.txPr!.sourceXml).toContain('sz="2400"');
	expect(chartExDataChanged(space, edited)).toBe(false);
});
