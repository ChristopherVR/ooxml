import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { applyChartExSpaceDataSheetEdits, chartExSpaceDataSheet, setChartExTitle } from '../chart';
import { readEmbeddedChartExSpace, writeEmbeddedChartExSpace } from './embedded-chart-ex-space';
import { loadXlsx } from './read';
import { saveXlsx } from './write';
import { createEditSession } from './edit/session';
import { readEmbeddedChart } from './ui/dialogs/grid-chart';
import { createVsdx } from '../visio/create-document';
import { editVsdx } from '../visio/edit';
import { parseVsdx } from '../visio/parser';
const fixture = () =>
	readFileSync(new URL('./__fixtures__/excel-chart-families.xlsx', import.meta.url));
async function embedded() {
	const book = await loadXlsx(fixture());
	const session = createEditSession(book);
	for (const [s, ws] of book.sheets.entries())
		for (let i = ws.drawings.length - 1; i >= 0; i--) {
			const d = ws.drawings[i]!;
			if (d.kind !== 'unsupported' || d.chartExSpace?.plotArea.series[0]?.layoutId !== 'waterfall')
				session.deleteDrawing(s, i);
		}
	const zip = await JSZip.loadAsync(await saveXlsx(book));
	zip.file('customXml/preservation.xml', '<keep exact="true"/>');
	const rels = 'xl/_rels/workbook.xml.rels';
	zip.file(
		rels,
		(await zip.file(rels)!.async('string')).replace(
			'</Relationships>',
			'<Relationship Id="preserved" Type="urn:test:provider" Target="../customXml/preservation.xml"/></Relationships>',
		),
	);
	return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
describe('embedded ChartEx source transactions', () => {
	it('admits ChartEx for Visio rendering and preserves a no-op workbook byte for byte', async () => {
		const bytes = await embedded();
		const found = (await readEmbeddedChartExSpace(bytes))!;
		expect(found.editable).toBe(true);
		expect(chartExSpaceDataSheet(found.space).categories).toHaveLength(6);
		expect(await writeEmbeddedChartExSpace(bytes, found.space)).toBe(bytes);
		expect((await readEmbeddedChart(bytes))!.view.type).toBe('waterfall');
		const result = await editVsdx(await createVsdx(), [
			{
				type: 'insert-chart-object',
				pageId: '0',
				shapeId: '1',
				x: 4,
				y: 5,
				width: 6,
				height: 4,
				workbook: bytes,
			},
		]);
		const back = (await parseVsdx(result.bytes)).pages[0]!.shapes[0]!.embeddedObject!;
		expect(back.chart.type).toBe('waterfall');
		expect(back.bytes).toEqual(bytes);
	});
	it('synchronizes edited source cells and keeps unrelated parts and the drawing', async () => {
		const bytes = await embedded();
		const found = (await readEmbeddedChartExSpace(bytes))!;
		const next = setChartExTitle(
			applyChartExSpaceDataSheetEdits(found.space, [
				{ kind: 'value', series: 0, category: 1, value: 42 },
			]),
			'Embedded waterfall',
		);
		const output = await writeEmbeddedChartExSpace(bytes, next);
		const back = (await readEmbeddedChartExSpace(output))!;
		expect(chartExSpaceDataSheet(back.space).series[0]!.values[1]).toBe(42);
		expect((await readEmbeddedChart(output))!.view.title).toBe('Embedded waterfall');
		expect(
			await (await JSZip.loadAsync(output)).file('customXml/preservation.xml')!.async('string'),
		).toBe('<keep exact="true"/>');
		expect(back.partName).toBe(found.partName);
	});
	it('renders live hidden-name data when Office omits the ChartEx caches', async () => {
		const bytes = await embedded();
		const found = (await readEmbeddedChartExSpace(bytes))!;
		const zip = await JSZip.loadAsync(bytes);
		zip.file(found.partName, found.sourceXml.replace(/<cx:lvl\b[^>]*>[\s\S]*?<\/cx:lvl>/gu, ''));
		const uncached = await zip.generateAsync({ type: 'uint8array' });
		const chart = (await readEmbeddedChart(uncached))!;
		expect(chart.view.categories).toHaveLength(6);
		expect(chart.view.series[0]!.values).toEqual(
			chartExSpaceDataSheet(found.space).series[0]!.values,
		);
	});
	it('refuses edits with unmodeled content while preserving an unchanged source', async () => {
		const bytes = await embedded();
		const found = (await readEmbeddedChartExSpace(bytes))!;
		const zip = await JSZip.loadAsync(bytes);
		zip.file(
			found.partName,
			found.sourceXml.replace('</cx:chartSpace>', '<cx:future/></cx:chartSpace>'),
		);
		const unknown = await zip.generateAsync({ type: 'uint8array' });
		const parsed = (await readEmbeddedChartExSpace(unknown))!;
		expect(parsed.editable).toBe(false);
		expect(await writeEmbeddedChartExSpace(unknown, parsed.space)).toBe(unknown);
		await expect(
			writeEmbeddedChartExSpace(unknown, setChartExTitle(parsed.space, 'Unsafe')),
		).rejects.toThrow(/preserv|unsupported/u);
	});
});
