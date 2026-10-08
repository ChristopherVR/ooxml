import { expect, test } from '@playwright/test';
import { createWorkbook, createEditSession, saveXlsx, loadXlsx } from 'ooxml-core/xlsx';
import type { DrawingFill } from 'ooxml-core/drawingml';
import { FRAMEWORKS, editor, openLanding, pageErrors } from './helpers';

for (const framework of FRAMEWORKS)
	test(`chart background fills survive UI export in ${framework}`, async ({ page }) => {
		const errors = pageErrors(page);
		await openLanding(page, framework);
		for (const kind of ['none', 'solid', 'linear', 'rect', 'circle', 'shape'] as const) {
			const fill: DrawingFill =
				kind === 'none'
					? { kind: 'none' }
					: kind === 'solid'
						? { kind: 'solid', color: { kind: 'srgb', value: 'F08020', transforms: [] } }
						: {
								kind: 'gradient',
								stops: [
									{ position: 0, color: { kind: 'srgb', value: 'FF0000', transforms: [] } },
									{ position: 100, color: { kind: 'srgb', value: 'FFFFFF', transforms: [] } },
								],
								...(kind === 'linear'
									? { angle: 90, scaled: true }
									: { path: kind, fillToRect: { l: 0.5, t: 0.5, r: 0.5, b: 0.5 } }),
							};
			const book = createWorkbook();
			createEditSession(book).addChart(0, {
				chartType: 'column',
				title: 'Fill export',
				showLegend: true,
				anchor: {
					from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 },
					ext: { cx: 600 * 9525, cy: 400 * 9525 },
				},
				series: [{ categories: ['A', 'B'], values: [10, 20] }],
				formatting: {
					sourceXml: '',
					entries: Object.fromEntries(
						['chartArea', 'plotArea', 'title', 'legend'].map((part) => [
							part,
							{ sourceXml: '', fill },
						]),
					),
				},
			});
			const name = `background-${kind}.xlsx`;
			await page.locator('#landing-file').setInputFiles({
				name,
				mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
				buffer: Buffer.from(await saveXlsx(book)),
			});
			const host = editor(page);
			await expect(host.getByText(name, { exact: true })).toBeVisible();
			await expect
				.poll(() =>
					host
						.locator('[data-chart-series]')
						.first()
						.evaluate(
							(node) =>
								node
									.closest('svg')!
									.querySelector(':scope > rect[x="0"][y="0"]')
									?.getAttribute('fill') ?? null,
						),
				)
				.toEqual(
					kind === 'none' ? null : kind === 'solid' ? '#F08020' : expect.stringMatching(/^url\(#/),
				);
			const bytes = await host.evaluate(async (node) =>
				Array.from(await (node as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
			);
			const back = await loadXlsx(Uint8Array.from(bytes));
			const chart = back.sheets[0]!.drawings.find((item) => item.kind === 'chart');
			if (chart?.kind !== 'chart') throw new Error('Missing exported chart');
			for (const part of ['chartArea', 'plotArea', 'title', 'legend']) {
				const actual = chart.formatting?.entries[part]?.fill;
				expect(actual?.kind).toBe(fill.kind);
				if (actual?.kind === 'gradient' && fill.kind === 'gradient') {
					expect(actual.stops).toEqual(fill.stops);
					expect(actual.path).toBe(fill.path);
					expect(actual.angle).toBe(fill.angle);
				} else expect(actual).toEqual(fill);
			}
		}
		expect(errors).toEqual([]);
	});
