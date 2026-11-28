// Bold text in a family without a bold face is emboldened, as Word does (issue #426).
import { inflateSync } from 'node:zlib';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import type { PdfFontFile } from '../../pdf/write/pdf-embedded-font';
import { syntheticTrueType } from '../../text/truetype/test-support/synthetic-truetype';
import { exportDocxPdf } from './export-pdf';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

async function pdfContent(body: string, fonts: PdfFontFile[]): Promise<string> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`,
	);
	const { model } = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	const { bytes } = await exportDocxPdf(model, { fonts });
	const raw = Buffer.from(bytes).toString('latin1');
	const streams: string[] = [];
	for (const match of raw.matchAll(/<<([^]*?)>>\nstream\n/g)) {
		const start = match.index + match[0].length;
		const length = Number(/\/Length (\d+)/.exec(match[1]!)?.[1]);
		if (/\/Filter \/FlateDecode/.test(match[1]!) && !/\/Subtype \/Image|\/Length1/.test(match[1]!))
			streams.push(inflateSync(bytes.subarray(start, start + length)).toString('latin1'));
	}
	return streams.join('\n');
}

const run = (props: string) =>
	`<w:p><w:r><w:rPr><w:rFonts w:ascii="Synth" w:hAnsi="Synth"/>${props}<w:sz w:val="32"/></w:rPr><w:t>AB</w:t></w:r></w:p>`;
const regular: PdfFontFile = { family: 'Synth', data: syntheticTrueType() };

describe('PDF export of bold text in a family without a bold face', () => {
	it('fills and strokes the regular face, keeping its advances', async () => {
		const content = await pdfContent(run('') + run('<w:b/>'), [regular]);
		const blocks = content.match(/BT [^]*? ET/g)!;
		expect(blocks).toHaveLength(2);
		expect(blocks[0]).not.toContain(' Tr');
		expect(blocks[1]).toContain(' 2 Tr');
		// The pen is the text colour and 1/24 em wide: 16 pt is 21.333 px, so 0.888889 px.
		expect(content).toMatch(/0 0 0 RG\n?\s*0\.888889 w/);
		// Same glyphs, same advances: only the render mode differs.
		const text = (block: string) =>
			block.replace(/ 2 Tr| 0 Tr/g, '').replace(/ -?[0-9.]+ [0-9.]+ Tm/, ' Tm');
		expect(text(blocks[1]!)).toBe(text(blocks[0]!));
	});

	it('does not embolden a face that is bold itself', async () => {
		const content = await pdfContent(run('<w:b/>'), [
			regular,
			{ family: 'Synth', bold: true, data: syntheticTrueType() },
		]);
		expect(content).not.toContain(' 2 Tr');
	});
});
