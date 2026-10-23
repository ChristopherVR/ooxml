import { describe, expect, it } from 'vitest';
import { parseChartSpace } from '../parse-space';
import { CHART_QUICK_LAYOUTS } from '../quick-layouts';
import { chartStylePresetOfColors, findChartStylePreset } from '../style-presets';
import { writeChartSpace } from '../write-space';
import { applyChartEditIntent } from './intent';
import {
	applyChartQuickLayout,
	chartQuickLayoutOf,
	chartSupportsQuickLayout,
} from './quick-layout';
import { applyChartStylePreset } from './style';
import { columnChart } from './test-support/column-chart';

const reload = (space: ReturnType<typeof columnChart>) => {
	const parsed = parseChartSpace(writeChartSpace(space));
	expect(parsed.issues).toEqual([]);
	return parsed.chartSpace;
};

describe('quick layout', () => {
	it('applies Layout 5: title, data table with keys, value axis title, no legend', () => {
		const space = reload(applyChartQuickLayout(columnChart(), 5));
		expect(space.title).toBeDefined();
		expect(space.legend).toBeUndefined();
		expect(space.plotArea.dataTable).toMatchObject({ showKeys: true, showHorizontalBorder: true });
		const [cat, val] = space.plotArea.axes;
		expect(val?.title?.text).toBe('Axis Title');
		expect(cat?.title).toBeUndefined();
		expect(val?.majorGridlines).toBe(true);
		expect(space.plotArea.groups[0]).toMatchObject({ gapWidth: 150, overlap: 0 });
		expect(chartQuickLayoutOf(space)).toBe(5);
	});

	it('applies Layout 2: top legend, outside-end value labels, hidden value axis', () => {
		const space = reload(applyChartQuickLayout(columnChart(), 2));
		expect(space.legend?.position).toBe('t');
		expect(space.plotArea.groups[0]?.dataLabels).toMatchObject({
			showValue: true,
			position: 'outEnd',
		});
		expect(space.plotArea.axes[1]).toMatchObject({ deleted: true, majorGridlines: false });
		expect(space.plotArea.dataTable).toBeUndefined();
		expect(space.plotArea.groups[0]?.overlap).toBe(-25);
	});

	it('recognises every layout it applies and keeps a rich title it already has', () => {
		for (const layout of CHART_QUICK_LAYOUTS) {
			const space = applyChartQuickLayout(columnChart(), layout.n);
			expect(chartQuickLayoutOf(space)).toBe(layout.n);
			if (layout.title) expect(space.title?.text).toBe('Sales 2026');
		}
		expect(chartQuickLayoutOf(columnChart())).toBeUndefined();
	});

	it('applies only the parts asked for', () => {
		const space = applyChartQuickLayout(columnChart(), 4, ['axes', 'dataTable']);
		expect(space.title?.text).toBe('Sales 2026');
		expect(space.plotArea.axes[1]?.majorGridlines).toBe(false);
		expect(space.plotArea.groups[0]?.gapWidth).toBe(219);
	});

	it('refuses unknown layouts and chart families without layouts', () => {
		expect(() => applyChartQuickLayout(columnChart(), 12)).toThrow(RangeError);
		const pie = columnChart();
		pie.plotArea.groups[0]!.kind = 'pie';
		expect(chartSupportsQuickLayout(pie)).toBe(false);
		expect(() => applyChartQuickLayout(pie, 1)).toThrow(RangeError);
	});

	it('runs as an edit intent', () => {
		const space = applyChartEditIntent(columnChart(), { op: 'quickLayout', layout: 7 });
		expect(space.plotArea.axes[1]?.minorGridlines).toBe(true);
	});
});

describe('style presets', () => {
	it('pins the preset colours per series and mirrors the style number', () => {
		const space = reload(applyChartStylePreset(columnChart(), 'pastel'));
		const colors = findChartStylePreset('pastel')!.colors();
		expect(space.style).toBe(42);
		const fills = space.plotArea.groups[0]!.series.map((series) =>
			series.spPr?.fill?.kind === 'solid' ? `#${series.spPr.fill.color.value}` : undefined,
		);
		expect(fills.map((c) => c?.toLowerCase())).toEqual(
			colors.slice(0, 2).map((c) => c.toLowerCase()),
		);
		expect(chartStylePresetOfColors(fills)?.id).toBe('pastel');
		expect(chartStylePresetOfColors(['#123456'])).toBeUndefined();
		expect(() => applyChartStylePreset(columnChart(), 'nope')).toThrow(RangeError);
	});

	it('runs style and colour intents', () => {
		expect(applyChartEditIntent(columnChart(), { op: 'style', id: 102 }).c14Style).toBe(102);
		expect(
			applyChartEditIntent(columnChart(), { op: 'stylePreset', preset: 'monochrome' }).style,
		).toBe(3);
		const recoloured = applyChartEditIntent(columnChart(), {
			op: 'colorPalette',
			from: 10,
			to: 12,
		});
		expect(recoloured).not.toBe(columnChart());
	});
});
