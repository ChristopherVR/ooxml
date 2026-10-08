import JSZip from 'jszip';
import { describe, it, expect } from 'vitest';

import { PresentationBuilder } from '../../core/builders/sdk/PresentationBuilder';
import { PptxHandler } from '../../core/PptxHandler';
import type { TablePptxElement } from '../../core/types/elements';

/** An empty cell sized only by its end paragraph properties, with no `a:tcPr`. */
const EMPTY_CELL =
	'<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>' +
	'<a:p><a:endParaRPr lang="en-US" sz="900"/></a:p>' +
	'</a:txBody></a:tc>';

const SLIDE = 'ppt/slides/slide1.xml';

async function buildDeck(): Promise<Uint8Array> {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	data.slides.push(
		createSlide('Blank')
			.addTable(
				{ rows: [{ cells: [{ text: 'Kept' }, { text: 'EmptyCellFixture' }] }] },
				{ x: 20, y: 20, width: 400, height: 80 },
			)
			.build(),
	);
	const zip = await JSZip.loadAsync(await handler.save(data.slides));
	const xml = await zip.file(SLIDE)!.async('string');
	const cell = /<a:tc>(?:(?!<\/a:tc>)[\s\S])*?<a:t>EmptyCellFixture<\/a:t>[\s\S]*?<\/a:tc>/;
	expect(xml, 'fixture precondition: marker cell').toMatch(cell);
	zip.file(SLIDE, xml.replace(cell, EMPTY_CELL));
	return zip.generateAsync({ type: 'uint8array' });
}

function emptyCellXml(xml: string): string {
	const match = xml.match(
		/<a:tc>(?:(?!<\/a:tc>)[\s\S])*?<a:endParaRPr lang="en-US" sz="900"[\s\S]*?<\/a:tc>/,
	);
	expect(match, 'empty cell in saved slide').not.toBeNull();
	return match![0];
}

describe('empty table cell without cell properties', () => {
	for (const moved of [false, true]) {
		it(`does not gain a:tcPr on an unedited save${moved ? ' of a moved table' : ''}`, async () => {
			const handler = new PptxHandler();
			const loaded = await handler.load((await buildDeck()).buffer as ArrayBuffer);
			const table = loaded.slides[0].elements.find(
				(element) => element.type === 'table',
			) as TablePptxElement;
			// The load sizes the cell from its end paragraph properties.
			expect(table.tableData!.rows[0].cells[1].style?.fontSize).toBe(9);
			if (moved) {
				// Move the table so the save writes it rather than passing it through.
				table.x += 1;
			}

			const saved = await handler.save(loaded.slides);
			const xml = await (await JSZip.loadAsync(saved)).file(SLIDE)!.async('string');
			const cell = emptyCellXml(xml);
			expect(cell).not.toMatch(/<a:tcPr\b/);
			expect(cell).toMatch(/<a:endParaRPr lang="en-US" sz="900"(?:\/>|>)/);
		});
	}
});
