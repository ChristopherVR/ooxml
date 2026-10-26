import { expect, it } from 'vitest';
import { createChartExSpace, applyChartExEditIntent } from 'ooxml-core/chart';
import { chartExSpaceFormatSource } from './format-pane-chartex';
it('exposes supported DrawingML controls with native axis/series controls left to ChartEx', () => {
	const space = createChartExSpace(
		{ categories: ['A'], series: [{ name: 'Sales', values: [1] }] },
		'waterfall',
		{ title: 'Title', legend: { visible: true, position: 'l' } },
	);
	const source = chartExSpaceFormatSource(space);
	expect(source.textOptions).toBe(true);
	expect(source.axisOptions).toBeUndefined();
	expect(source.seriesOptions).toBeUndefined();
	expect(source.element({ kind: 'title' }).support).toMatchObject({
		pattern: true,
		effects: true,
		text: true,
		lineDetails: true,
		secondaryAxis: false,
	});
	expect(source.legendOptions?.()?.position).toBe('l');
	const edited = applyChartExEditIntent(space, {
		op: 'text',
		target: { kind: 'title' },
		edit: { sizePt: 24, bold: true },
	});
	expect(edited.title!.tx!.rich!.sourceXml).toContain('sz="2400"');
	expect(() =>
		applyChartExEditIntent(space, { op: 'secondaryAxis', seriesIndex: 0, secondary: true }),
	).toThrow('does not support');
});
