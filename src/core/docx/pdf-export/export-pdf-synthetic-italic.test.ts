import { inflateSync } from 'node:zlib';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { createFakeMeasurer } from '../layout/index';
import { syntheticTrueType } from '../../text/truetype/test-support/synthetic-truetype';
import { exportDocxPdf } from './export-pdf';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

async function pdfContent(body: string): Promise<string> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`,
	);
	const { model } = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	const { bytes } = await exportDocxPdf(model, {
		measurer: createFakeMeasurer(),
		fonts: [{ family: 'Synth', data: syntheticTrueType() }],
	});
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
	`<w:p><w:r><w:rPr><w:rFonts w:ascii="Synth" w:hAnsi="Synth"/>${props}</w:rPr><w:t>AB</w:t></w:r></w:p>`;

describe('PDF export of italic text in a family without an italic face', () => {
	it('shears the regular face, as Word does, and leaves regular text upright', async () => {
		const content = await pdfContent(run('') + run('<w:i/>') + run('<w:b/><w:i/>'));
		const matrices = [...content.matchAll(/1 0 ([\d.]+) -1 [\d.]+ [\d.]+ Tm/g)].map((m) => m[1]);
		// Regular text stays upright; italic and bold italic lean by Word's slope.
		expect(matrices).toEqual(['0', '0.333333', '0.333333']);
		// The text itself is unchanged.
		expect(content.match(/<00010002>/g)).toHaveLength(3);
	});
});
