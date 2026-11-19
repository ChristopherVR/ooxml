import { strToU8, zipSync } from 'fflate';
import JSZip from 'jszip';

import { docx, p } from '../../../docx/test-support/compare-docs';
import { loadDocx } from '../../../docx/parse';
import { readDeclaredFonts } from '../../../opc/declared-fonts';
import { PresentationBuilder } from '../../../pptx/core/builders/sdk';
import { PptxHandler } from '../../../pptx/core/PptxHandler';
import { loadXlsx } from '../../../xlsx/read/load';
import type { SecurityCase } from '../case';

// Each input repeats an element that a text scan took for an opening tag and then searched to
// the end of the part for its closing tag: time grew with the square of the repetitions.

const S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PR = 'http://schemas.openxmlformats.org/package/2006/relationships';

/** A one-sheet workbook whose header and footer drawing is `vml`. */
async function workbookWithHeaderDrawing(vml: string): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
			'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
			'<Default Extension="xml" ContentType="application/xml"/>' +
			'<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>' +
			'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
			'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
	);
	const rels = (id: string, type: string, target: string) =>
		`<Relationships xmlns="${PR}"><Relationship Id="${id}" Type="${R}/${type}" Target="${target}"/></Relationships>`;
	zip.file('_rels/.rels', rels('rId1', 'officeDocument', 'xl/workbook.xml'));
	zip.file(
		'xl/workbook.xml',
		`<workbook xmlns="${S}" xmlns:r="${R}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
	);
	zip.file('xl/_rels/workbook.xml.rels', rels('rId1', 'worksheet', 'worksheets/sheet1.xml'));
	zip.file(
		'xl/worksheets/sheet1.xml',
		`<worksheet xmlns="${S}" xmlns:r="${R}"><sheetData/><legacyDrawingHF r:id="rId9"/></worksheet>`,
	);
	zip.file('xl/worksheets/_rels/sheet1.xml.rels', rels('rId9', 'vmlDrawing', '../drawings/v.vml'));
	zip.file('xl/drawings/v.vml', vml);
	return zip.generateAsync({ type: 'uint8array' });
}

export const cases: readonly SecurityCase[] = [
	{
		name: 'font declarations that never close are read in linear time',
		limits: { ms: 1_500, heapMb: 128 },
		run() {
			const count = 40_000;
			const result = readDeclaredFonts(
				zipSync({
					'word/fontTable.xml': strToU8(
						`<w:fonts>${'<w:font w:name="a"/>'.repeat(count)}</w:fonts>`,
					),
					'visio/document.xml': strToU8('<FaceName '.repeat(count)),
				}),
			);
			return { families: result.families };
		},
		expected: { families: ['a'] },
	},
	{
		name: 'a header drawing of empty shapes loads in linear time',
		limits: { ms: 2_000, heapMb: 192 },
		async run() {
			const vml = `<xml xmlns:v="urn:schemas-microsoft-com:vml">${'<v:shape/>'.repeat(80_000)}</xml>`;
			const book = await loadXlsx(await workbookWithHeaderDrawing(vml));
			return { sheets: book.sheets.length };
		},
		expected: { sheets: 1 },
	},
	{
		name: 'a long chain of based-on styles loads in linear time',
		limits: { ms: 6_000, heapMb: 512 },
		async run() {
			const count = 14_000;
			const styles = Array.from(
				{ length: count },
				(_, index) =>
					`<w:style w:type="paragraph" w:styleId="s${index}">` +
					`${index ? `<w:basedOn w:val="s${index - 1}"/>` : ''}<w:rPr><w:b/></w:rPr></w:style>`,
			).join('');
			const zip = await JSZip.loadAsync(await docx(p('text')));
			zip.file(
				'word/styles.xml',
				`<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${styles}</w:styles>`,
			);
			const document = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
			return { loaded: document !== undefined };
		},
		expected: { loaded: true },
	},
	{
		name: 'a slide of empty custom geometries loads in linear time',
		limits: { ms: 4_000, heapMb: 256 },
		async run() {
			const { handler, data, createSlide } = await PresentationBuilder.create();
			data.slides = [createSlide('Blank').addText('Kept text').build()];
			const zip = await JSZip.loadAsync(await handler.save(data.slides));
			handler.dispose();
			const slide = Object.keys(zip.files).find((path) => /^ppt\/slides\/[^/]+\.xml$/u.test(path))!;
			const xml = await zip.file(slide)!.async('string');
			zip.file(slide, xml.replace('</p:spTree>', `${'<a:custGeom/>'.repeat(120_000)}</p:spTree>`));
			const reader = new PptxHandler();
			const loaded = await reader.load(
				(await zip.generateAsync({ type: 'uint8array' })).slice().buffer,
			);
			reader.dispose();
			return { slides: loaded.slides.length };
		},
		expected: { slides: 1 },
	},
];
