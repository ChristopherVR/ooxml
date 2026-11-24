import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { simplePdf, stream } from '../../pdf/test-support/build';
import { importPdf } from './index';

const fixtures = JSON.parse(
	readFileSync(new URL('../../pdf/jpx/__fixtures__/jpx.json', import.meta.url), 'utf8'),
) as { width: number; height: number; cases: { name: string; data: string }[] };
const jpx = (name: string) =>
	new Uint8Array(Buffer.from(fixtures.cases.find((item) => item.name === name)!.data, 'base64'));

describe('importPdf with JPEG 2000 pictures', () => {
	it('converts JPXDecode images into pictures', async () => {
		const pdf = simplePdf(
			'BT /F1 12 Tf 72 700 Td (Caption) Tj ET q 200 0 0 100 72 500 cm /Im Do Q',
			{
				resources: '/XObject << /Im 6 0 R >>',
				objects: [
					stream(
						`/Type /XObject /Subtype /Image /Width ${fixtures.width} /Height ${fixtures.height} /Filter /JPXDecode`,
						jpx('rgb-97'),
						true,
					),
				],
			},
		);
		const { model, report, document } = await importPdf(pdf);
		const pictures = model.blocks.flatMap((block) =>
			block.type === 'paragraph' ? block.runs.filter((run) => run.image) : [],
		);
		expect(pictures).toHaveLength(1);
		expect(document.media?.size).toBe(1);
		expect(report.pages[0]!.notes.filter((note) => note.includes('picture was left out'))).toEqual(
			[],
		);
	});

	it('reports a JPEG 2000 picture it cannot decode', async () => {
		const pdf = simplePdf('q 200 0 0 100 72 500 cm /Im Do Q', {
			resources: '/XObject << /Im 6 0 R >>',
			objects: [
				stream(
					'/Type /XObject /Subtype /Image /Width 8 /Height 8 /Filter /JPXDecode',
					'damaged',
					true,
				),
			],
		});
		const { report } = await importPdf(pdf);
		expect(report.pages[0]!.notes).toContain('A picture was left out: it is not a JPEG 2000 file.');
	});
});
