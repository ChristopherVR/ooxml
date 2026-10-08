import JSZip from 'jszip';
import { describe, it, expect } from 'vitest';

import { PresentationBuilder } from '../../core/builders/sdk/PresentationBuilder';
import { PptxHandler } from '../../core/PptxHandler';
import type { TablePptxElement } from '../../core/types/elements';

/** An empty first paragraph whose end properties are smaller than the text below it. */
const CELL_BODY =
	'<a:txBody><a:bodyPr/><a:lstStyle/>' +
	'<a:p><a:endParaRPr lang="en-US" sz="600"/></a:p>' +
	'<a:p><a:r><a:rPr lang="en-US" sz="1200"/><a:t>Below</a:t></a:r></a:p>' +
	'</a:txBody>';

async function buildDeck(): Promise<Uint8Array> {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	data.slides.push(
		createSlide('Blank')
			.addTable(
				{ rows: [{ cells: [{ text: 'EndParaFixture' }, { text: '' }] }] },
				{ x: 20, y: 20, width: 400, height: 80 },
			)
			.build(),
	);
	const zip = await JSZip.loadAsync(await handler.save(data.slides));
	const path = 'ppt/slides/slide1.xml';
	const xml = await zip.file(path)!.async('string');
	const body =
		/<a:txBody>(?:(?!<\/a:txBody>)[\s\S])*?<a:t>EndParaFixture<\/a:t>(?:(?!<\/a:txBody>)[\s\S])*?<\/a:txBody>/;
	expect(xml, 'fixture precondition: marker body').toMatch(body);
	zip.file(path, xml.replace(body, CELL_BODY));
	return zip.generateAsync({ type: 'uint8array' });
}

describe('table cell end paragraph size on save', () => {
	it('keeps the text run size of a cell whose first paragraph is empty', async () => {
		const fixture = await buildDeck();
		const sourceXml = await (
			await JSZip.loadAsync(fixture)
		)
			.file('ppt/slides/slide1.xml')!
			.async('string');
		const handler = new PptxHandler();
		const loaded = await handler.load(fixture.buffer as ArrayBuffer);
		const table = loaded.slides[0].elements.find(
			(element) => element.type === 'table',
		) as TablePptxElement;
		expect(table.tableData!.rows[0].cells[0].style?.fontSize).toBeUndefined();
		// Move the table so the save writes it rather than passing it through.
		table.x += 1;

		const saved = await handler.save(loaded.slides);
		const xml = await (await JSZip.loadAsync(saved)).file('ppt/slides/slide1.xml')!.async('string');
		expect(xml).toMatch(/<a:rPr lang="en-US" sz="1200"(?:\/>|><\/a:rPr>)<a:t>Below<\/a:t>/);
		expect(xml).toMatch(/<a:endParaRPr lang="en-US" sz="600"(?:\/>|>)/);
		// No cell properties were added for the cell.
		const tcPrCount = (text: string) => text.match(/<a:tcPr\b/g)?.length ?? 0;
		expect(tcPrCount(xml)).toBe(tcPrCount(sourceXml));
	});
});
