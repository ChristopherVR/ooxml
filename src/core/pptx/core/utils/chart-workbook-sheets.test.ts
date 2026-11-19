import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { chartPartSheets, placeChartData, type ChartDataSheets } from './chart-workbook-sheets';

const rels = (target: string) =>
	`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="${target}"/></Relationships>`;
const part = (...formulas: string[]) =>
	`<c:chartSpace xmlns:c="c">${formulas.map((f) => `<c:f>${f}</c:f>`).join('')}</c:chartSpace>`;

function deck(parts: Record<string, string>) {
	const zip = new JSZip();
	for (const [name, xml] of Object.entries(parts)) {
		zip.file(`ppt/charts/${name}.xml`, xml);
		zip.file(`ppt/charts/_rels/${name}.xml.rels`, rels('../embeddings/book.xlsx'));
	}
	return zip;
}
const session = (allocated: string[], held: string[] = []): ChartDataSheets => ({
	allocated: new Map([['ppt/embeddings/book.xlsx', new Set(allocated)]]),
	held: () => new Set(held),
});

describe('chartPartSheets', () => {
	it('reads every sheet the formulas of a part name, quoted or not', () => {
		const xml = `${part("'It''s &amp; more'!$A$1", '(Sheet1!$A$2,Data!$B$2:$B$3)')}<cx:f dir="row">ChartExData1!$A$1</cx:f><a:fld/>`;
		expect([...chartPartSheets(xml)]).toEqual(["it's & more", 'sheet1', 'data', 'chartexdata1']);
	});
});

describe('placeChartData', () => {
	it('writes over the sheet it allocated that the chart reads and removes the ones nothing reads', async () => {
		const zip = deck({ chart1: part('ChartData2!$A$1') });
		const sheets = session(['chartdata1', 'chartdata2']);
		const place = await placeChartData(
			zip,
			'ppt/embeddings/book.xlsx',
			'ppt/charts/chart1.xml',
			sheets,
		);
		expect(place.options).toEqual({ reuse: 'chartdata2', drop: ['chartdata1', 'chartdata2'] });
		place.written('ChartData2');
		expect([...sheets.allocated.get('ppt/embeddings/book.xlsx')!]).toEqual(['chartdata2']);
	});

	it('never names a sheet it did not allocate, or one another chart or a saved state reads', async () => {
		const zip = deck({
			chart1: part('Sheet1!$A$1', 'ChartData1!$A$1'),
			chart2: part('ChartData1!$A$2'),
		});
		const sheets = session(['chartdata1', 'chartdata2', 'chartdata3'], ['chartdata3']);
		const place = await placeChartData(
			zip,
			'ppt/embeddings/book.xlsx',
			'ppt/charts/chart1.xml',
			sheets,
		);
		expect(place.options).toEqual({ drop: ['chartdata2'] });
		place.written('ChartData4');
		expect([...sheets.allocated.get('ppt/embeddings/book.xlsx')!]).toEqual([
			'chartdata4',
			'chartdata1',
			'chartdata3',
		]);
	});
});
