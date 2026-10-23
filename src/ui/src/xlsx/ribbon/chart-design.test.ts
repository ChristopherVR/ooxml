// @vitest-environment jsdom
import { afterEach, expect, it } from 'vitest';
import { createWorkbook, loadXlsx, saveXlsx, type ChartObject } from 'ooxml-core/xlsx';
import { contextualCommands } from 'ooxml-core/xlsx/ui';
import { createTestContext } from '../commands/test-support';
import {
	chartColorGalleryItems,
	chartColorGallerySections,
	chartQuickLayoutGalleryItems,
	chartQuickLayoutGallerySections,
	chartStyleGalleryItems,
} from '../commands/chart-design';
import { renderSharedGallery } from './shared-gallery';

afterEach(() => {
	document.body.innerHTML = '';
});

it('uses the shared gallery to recolor a selected chart, save/reload and undo', async () => {
	const ctx = createTestContext(createWorkbook());
	expect(chartColorGallerySections(ctx)).toEqual([]);
	ctx.session()!.addChart(0, {
		chartType: 'column',
		anchor: {
			from: { row: 0, col: 3, colOffset: 0, rowOffset: 0 },
			ext: { cx: 4000000, cy: 3000000 },
		},
		showLegend: true,
		series: [
			{ name: 'Sales', categories: ['a', 'b'], values: [2, 4] },
			{ name: 'Cost', categories: ['a', 'b'], values: [1, 3] },
		],
	});
	ctx.selection.set({ drawing: 0 });
	ctx.commands.registerAll(contextualCommands());
	const command = ctx.commands.get('chart.colors')!;
	const updates: (() => void)[] = [];
	const gallery = renderSharedGallery(
		{ ctx, doc: document, updates, isHidden: () => false },
		{
			kind: 'gallery',
			command: command.id,
			items: chartColorGalleryItems,
			sections: chartColorGallerySections,
		},
		command,
	);
	ctx.root.append(gallery);
	for (const update of updates) update();
	const trigger = gallery.querySelector<HTMLButtonElement>('[data-gallery="chart.colors"]')!;
	expect(trigger.dataset.command).toBe('chart.colors');
	trigger.click();
	expect(gallery.querySelectorAll('.popup .section')).toHaveLength(2);
	expect(gallery.querySelectorAll('.popup .tile svg')).toHaveLength(17);
	gallery.querySelector<HTMLButtonElement>('[data-gallery-item="15"]')!.click();
	await Promise.resolve();
	expect((ctx.workbook()!.sheets[0]!.drawings[0] as ChartObject).colorPalette).toBe(15);
	for (const update of updates) update();
	trigger.click();
	expect(gallery.querySelector('[data-gallery-item="15"]')?.getAttribute('aria-pressed')).toBe(
		'true',
	);
	const translate = ctx.t;
	ctx.t = (key, vars) => `Translated ${translate(key, vars)}`;
	for (const update of updates) update();
	expect(trigger.getAttribute('aria-label')).toBe('Translated Change Colors');
	expect(gallery.querySelector('[data-gallery-item="15"]')?.getAttribute('aria-label')).toBe(
		'Translated Monochromatic Palette 2',
	);
	const loaded = await loadXlsx(await saveXlsx(ctx.workbook()!));
	expect((loaded.sheets[0]!.drawings[0] as ChartObject).colorPalette).toBe(15);
	ctx.session()!.undo();
	expect((ctx.workbook()!.sheets[0]!.drawings[0] as ChartObject).colorPalette).toBeUndefined();
	ctx.setReadOnly(true);
	for (const update of updates) update();
	expect(gallery.querySelector<HTMLButtonElement>('.trigger')!.disabled).toBe(true);
});

it('applies the shared Chart Styles and Quick Layout galleries through the ChartPatch path', async () => {
	const ctx = createTestContext(createWorkbook());
	ctx.session()!.addChart(0, {
		chartType: 'column',
		anchor: {
			from: { row: 0, col: 3, colOffset: 0, rowOffset: 0 },
			ext: { cx: 4000000, cy: 3000000 },
		},
		showLegend: true,
		series: [
			{ name: 'Sales', categories: ['a', 'b'], values: [2, 4] },
			{ name: 'Cost', categories: ['a', 'b'], values: [1, 3] },
		],
	});
	ctx.selection.set({ drawing: 0 });
	ctx.commands.registerAll(contextualCommands());
	const chart = () => ctx.workbook()!.sheets[0]!.drawings[0] as ChartObject;
	expect(chartStyleGalleryItems(ctx).map((item) => item.id)).toEqual([
		'colorful',
		'monochrome',
		'colorfulLight',
		'colorfulDark',
		'mutedDark',
		'pastel',
	]);
	await ctx.commands.run('chart.styles', 'monochrome');
	expect(chartStyleGalleryItems(ctx).find((item) => item.applied)?.id).toBe('monochrome');
	const layouts = chartQuickLayoutGallerySections(ctx);
	expect(layouts[0]?.items).toHaveLength(11);
	expect(layouts[0]?.items[4]?.label).toBe('Layout 5');
	await ctx.commands.run('chart.quick-layout', 'layout4');
	expect(chart()).toMatchObject({ quickLayout: 4, showLegend: true, legendPosition: 'b' });
	expect(chart().title).toBeUndefined();
	expect(chartQuickLayoutGalleryItems(ctx).find((item) => item.applied)?.id).toBe('layout4');
	const loaded = await loadXlsx(await saveXlsx(ctx.workbook()!));
	const back = loaded.sheets[0]!.drawings[0] as ChartObject;
	expect(back.series[0]?.dataLabels?.chart?.showValue).toBe(true);
	ctx.session()!.undo();
	expect(chart().quickLayout).toBeUndefined();
	ctx.session()!.updateChart(0, 0, { chartType: 'pie' });
	expect(chartQuickLayoutGallerySections(ctx)).toEqual([]);
	expect(ctx.commands.isEnabled('chart.quick-layout')).toBe(false);
});
