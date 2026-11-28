import { inflateSync } from 'node:zlib';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { createFakeMeasurer } from '../layout/index';
import { exportDocxPdf } from './export-pdf';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const SHADE = '0.85098 0.882353 0.952941 rg';

const cell = (inner: string, props = '') =>
	`<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/>${props}</w:tcPr>${inner}</w:tc>`;
const text = (value: string) => `<w:p><w:r><w:t>${value}</w:t></w:r></w:p>`;
const shd = '<w:shd w:val="clear" w:color="auto" w:fill="D9E1F3"/>';

describe('PDF export of a shaded vertical merge', () => {
	it('paints every cell shading before any cell text', async () => {
		const body = `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
<w:tr>${cell(text('A1'))}${cell('<w:p/><w:p/><w:p/>' + text('Merged text'), `<w:vMerge w:val="restart"/>${shd}`)}</w:tr>
<w:tr>${cell(text('A2') + text('A3') + text('A4'))}${cell('<w:p/>', `<w:vMerge/>${shd}`)}</w:tr>
<w:tr>${cell(text('A5'))}${cell('<w:p/>', `<w:vMerge/>${shd}`)}</w:tr></w:tbl><w:p/>`;
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`,
		);
		const { model } = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const { bytes } = await exportDocxPdf(model, { measurer: createFakeMeasurer() });
		const raw = Buffer.from(bytes).toString('latin1');
		const streams: string[] = [];
		for (const match of raw.matchAll(/<<([^]*?)>>\nstream\n/g)) {
			const start = match.index + match[0].length;
			const length = Number(/\/Length (\d+)/.exec(match[1]!)?.[1]);
			if (/\/Filter \/FlateDecode/.test(match[1]!) && !/\/Subtype \/Image/.test(match[1]!))
				streams.push(inflateSync(bytes.subarray(start, start + length)).toString('latin1'));
		}
		const content = streams.join('\n');
		const lastShading = content.lastIndexOf(SHADE);
		const firstText = content.indexOf('BT');
		expect(lastShading).toBeGreaterThan(-1);
		expect(firstText).toBeGreaterThan(-1);
		expect(lastShading).toBeLessThan(firstText);
		expect(content.split(SHADE).length - 1).toBeGreaterThanOrEqual(3);
		expect(content).toContain('(Merged');
	});
});
