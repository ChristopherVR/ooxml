import { readFile } from 'node:fs/promises';
import { inflateSync } from 'node:zlib';
import { expect, it } from 'vitest';
import { loadDocx } from '../index';
import { createFakeMeasurer } from '../layout/index';
import { exportDocxPdf } from './export-pdf';

async function content(name: string): Promise<string> {
	const { model } = await loadDocx(
		await readFile(new URL(`../layout/fixtures/word-layout-reports/${name}.docx`, import.meta.url)),
	);
	const { bytes } = await exportDocxPdf(model, { measurer: createFakeMeasurer() });
	const raw = Buffer.from(bytes).toString('latin1');
	const streams: string[] = [];
	for (const match of raw.matchAll(/<<([^]*?)>>\nstream\n/g)) {
		const start = match.index + match[0].length;
		const length = Number(/\/Length (\d+)/.exec(match[1]!)?.[1]);
		if (/\/Filter \/FlateDecode/.test(match[1]!) && !/\/Subtype \/Image/.test(match[1]!))
			streams.push(inflateSync(bytes.subarray(start, start + length)).toString('latin1'));
	}
	return streams.join('\n');
}

// Word 16 draws automatic-colour text white on a dark shading and keeps an explicit black (#436).
// The fixture has six white texts of two words (cells 000000, 371127, C00000 and paragraphs of the same
// three fills) and black ones on 404040, 595959 and 4472C4.
it('exports automatic text on dark shading in white (#436)', async () => {
	const pdf = await content('issue-436-auto-colour-on-dark-shading');
	const white = /1 1 1 rg\s*BT/g;
	expect(pdf.match(white)?.length ?? 0).toBe(12);
});

it('exports explicit black on dark shading in black (#436 control)', async () => {
	const pdf = await content('issue-436-auto-colour-on-dark-shading-control');
	expect(pdf.match(/1 1 1 rg\s*BT/g)).toBeNull();
});
