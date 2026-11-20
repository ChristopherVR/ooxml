import { readFileSync } from 'node:fs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { loadDocx, type InlineImage } from './index';
import {
	CHART_PARTS,
	cellFloatCases,
	cellFloatDemo,
	cellFloatPackage,
	type CellFloatCase,
} from './__fixtures__/cell-floats/cases';

const chartPackage = unzipSync(readFileSync(new URL('./__fixtures__/chart.docx', import.meta.url)));
const chartParts = Object.fromEntries(
	CHART_PARTS.map((part) => [part, strFromU8(chartPackage[part]!)]),
);
const build = (entry: CellFloatCase) =>
	zipSync(
		Object.fromEntries(
			Object.entries(cellFloatPackage(entry, chartParts)).map(([path, xml]) => [
				path,
				strToU8(xml),
			]),
		),
	);
const documentXml = (bytes: Uint8Array) => strFromU8(unzipSync(bytes)['word/document.xml']!);

function images(model: Awaited<ReturnType<typeof loadDocx>>['model']): InlineImage[] {
	const found: InlineImage[] = [];
	const visit = (blocks: typeof model.blocks) => {
		for (const block of blocks)
			if (block.type === 'paragraph')
				for (const run of block.runs) {
					if (run.image) found.push(run.image);
				}
			else for (const row of block.rows) for (const cell of row) visit(cell.paragraphs);
	};
	visit(model.blocks);
	return found;
}

describe('floating objects in cells and floating charts through saving', () => {
	const entries: [string, CellFloatCase][] = [
		...Object.entries(cellFloatCases()),
		['demo', cellFloatDemo()],
	];

	it.each(entries)('saves %s unchanged byte for byte', async (_name, entry) => {
		const bytes = build(entry);
		const loaded = await loadDocx(bytes);
		expect(documentXml(await loaded.save(loaded.model))).toBe(documentXml(bytes));
	});

	it('reads a floating chart with its placement and layoutInCell="0"', async () => {
		const chart = images((await loadDocx(build(cellFloatCases()['chart-body']!))).model).find(
			(image) => image.chart,
		)!;
		expect(chart.anchored).toBe(true);
		expect(chart.placement).toMatchObject({
			wrap: 'square',
			relativeFrom: 'margin',
			offsetXPx: 48,
			relativeFromV: 'paragraph',
			offsetYPx: 8,
		});
		expect(chart.placement?.layoutInCell).toBeUndefined();
		const shape = images((await loadDocx(build(cellFloatCases()['h-page-lic0-none']!))).model)[0]!;
		expect(shape.placement?.layoutInCell).toBe(false);
	});

	it('saves a floating chart moved on the page, keeping its chart and wrapping', async () => {
		const loaded = await loadDocx(build(cellFloatCases()['chart-body']!));
		const chart = images(loaded.model).find((image) => image.chart)!;
		chart.placement = { ...chart.placement!, offsetXPx: 96, offsetYPx: 16 };
		const xml = documentXml(await loaded.save(loaded.model));
		expect(xml).toContain(
			`<wp:positionH relativeFrom="margin"><wp:posOffset>${96 * 9525}</wp:posOffset></wp:positionH>`,
		);
		expect(xml).toContain(
			`<wp:positionV relativeFrom="paragraph"><wp:posOffset>${16 * 9525}</wp:posOffset></wp:positionV>`,
		);
		expect(xml).toContain('<c:chart r:id="rIdChart"/>');
		expect(xml).toContain('<wp:wrapSquare wrapText="bothSides"/>');
	});
});
