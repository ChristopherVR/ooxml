// Fonts embedded in the document (`w:embedRegular`, `word/fonts/*.odttf`) in the PDF export (#418).
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { obfuscateFont } from '../../opc/font-obfuscation';
import { syntheticTrueType } from '../../text/truetype/test-support/synthetic-truetype';
import { exportDocxPdf } from './export-pdf';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const KEY = '{3A5C1F27-9B4E-4D80-A1C6-72E5D0B8F419}';

// 200 letters A: the synthetic face advances 0.6 em per A, so a 12 pt line holds 65 of them.
const body = `<w:p><w:r><w:rPr><w:rFonts w:ascii="Chivo" w:hAnsi="Chivo"/><w:sz w:val="24"/></w:rPr><w:t>${'A'.repeat(200)}</w:t></w:r></w:p>`;

async function docx(embed: boolean): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
	);
	if (embed) {
		zip.file(
			'word/fontTable.xml',
			`<w:fonts xmlns:w="${W}" xmlns:r="${R}"><w:font w:name="Chivo"><w:charset w:val="00"/><w:family w:val="auto"/><w:pitch w:val="variable"/><w:embedRegular r:id="rIdF1" w:fontKey="${KEY}"/></w:font></w:fonts>`,
		);
		zip.file(
			'word/_rels/fontTable.xml.rels',
			`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdF1" Type="${R}/font" Target="fonts/font1.odttf"/></Relationships>`,
		);
		zip.file('word/fonts/font1.odttf', obfuscateFont(syntheticTrueType(), KEY));
	}
	return zip.generateAsync({ type: 'uint8array' });
}

describe('exportDocxPdf embedded fonts', () => {
	it('measures and draws with the faces the document embeds', async () => {
		const { model } = await loadDocx(await docx(true));
		expect(model.embeddedFonts).toEqual([
			expect.objectContaining({ family: 'Chivo', bold: false, italic: false }),
		]);
		const { bytes, report } = await exportDocxPdf(model);
		const raw = Buffer.from(bytes).toString('latin1');
		expect(report.embeddedFonts).toEqual(['Chivo']);
		expect(raw).toMatch(/\/FontFile2/);
		expect(report.fontMeasurement).toEqual([
			{ family: 'Chivo', font: 'Chivo', source: 'supplied' },
		]);
	});

	it('leaves a family the host supplies to the host', async () => {
		const { model } = await loadDocx(await docx(true));
		const host = syntheticTrueType([], 1, 0);
		const { report } = await exportDocxPdf(model, { fonts: [{ family: 'Chivo', data: host }] });
		expect(report.embeddedFonts).toEqual(['Chivo']);
	});

	it('estimates the family without the embedded parts (control)', async () => {
		const { model } = await loadDocx(await docx(false));
		const { bytes, report } = await exportDocxPdf(model);
		expect(Buffer.from(bytes).toString('latin1')).not.toMatch(/\/FontFile2/);
		expect(report.embeddedFonts).toEqual([]);
	});
});
