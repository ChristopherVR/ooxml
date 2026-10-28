import { readFileSync } from 'node:fs';
import path from 'node:path';
import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import { loadXlsx, saveXlsx } from 'ooxml-core/xlsx';
import { FRAMEWORKS, editor, openLanding } from './helpers';

const points = (values: (number | string)[]) =>
	`<cx:lvl ptCount="${values.length}">${values.map((v, i) => `<cx:pt idx="${i}">${v}</cx:pt>`).join('')}</cx:lvl>`;
const chartXml = (kind: string) =>
	`<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chartData><cx:data id="0"><cx:strDim type="cat">${points(['A', 'A', 'A', 'B', 'B', 'B'])}</cx:strDim><cx:numDim type="val">${points([1, 2, 3, 4, 5, 6])}</cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="${kind}"><cx:dataId val="0"/></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>`;
async function fixture(kind: string) {
	const book = await loadXlsx(
		readFileSync(
			path.join(import.meta.dirname, '../../src/core/xlsx/__fixtures__/excel-chart-families.xlsx'),
		),
	);
	const drawing = book.sheets[0]!.drawings.find(
		(d) => d.kind === 'unsupported' && d.chartExSpace?.plotArea.series[0]?.layoutId === 'waterfall',
	);
	if (drawing?.kind !== 'unsupported' || !drawing.chartExPartName)
		throw new Error('Missing ChartEx source host');
	drawing.anchor = {
		from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 },
		ext: { cx: 600 * 9525, cy: 350 * 9525 },
	};
	book.sheets[0]!.drawings = [drawing];
	const zip = await JSZip.loadAsync(await saveXlsx(book));
	const xml = chartXml(kind);
	zip.file(drawing.chartExPartName, xml);
	return {
		part: drawing.chartExPartName,
		xml,
		bytes: await zip.generateAsync({ type: 'uint8array' }),
	};
}
for (const framework of FRAMEWORKS)
	for (const kind of ['funnel', 'boxWhisker'])
		test(`${framework}: shared ${kind} paints in Excel and saves its source unchanged`, async ({
			page,
		}, testInfo) => {
			const source = await fixture(kind);
			await page.setViewportSize({ width: 1400, height: 900 });
			await openLanding(page, framework);
			await page.locator('#landing-file').setInputFiles({
				name: 'neutral-chart.xlsx',
				mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
				buffer: Buffer.from(source.bytes),
			});
			const host = editor(page);
			const chart = host.locator('.xg-obj[data-index="0"]');
			await expect(
				chart.locator(kind === 'boxWhisker' ? '[data-chart-part="box"]' : 'g[data-chart-point]'),
			).toHaveCount(kind === 'boxWhisker' ? 2 : 6);
			await testInfo.attach(`${kind}-browser.png`, {
				body: await chart.screenshot(),
				contentType: 'image/png',
			});
			const bytes = Uint8Array.from(
				await host.evaluate(async (element) =>
					Array.from(
						await (element as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes(),
					),
				),
			);
			const saved = await JSZip.loadAsync(bytes);
			expect(await saved.file(source.part)!.async('string')).toBe(source.xml);
		});
