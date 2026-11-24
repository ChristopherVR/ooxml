import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { simplePdf, stream } from '../../pdf/test-support/build';
import { importPdf } from './index';

const fixture = (
	JSON.parse(
		readFileSync(new URL('../../pdf/__fixtures__/jbig2/jbig2.json', import.meta.url), 'utf8'),
	) as {
		name: string;
		globals?: string;
		data: string;
		pages: { width: number; height: number }[];
	}[]
).find((item) => item.name === 'symbol')!;
const bytes = (base64: string) => new Uint8Array(Buffer.from(base64, 'base64'));

describe('importPdf with JBIG2 pictures', () => {
	it('converts JBIG2Decode images (with globals) into pictures', async () => {
		const { width, height } = fixture.pages[0]!;
		const pdf = simplePdf('BT /F1 12 Tf 72 700 Td (Scan) Tj ET q 240 0 0 90 72 500 cm /Im Do Q', {
			resources: '/XObject << /Im 6 0 R >>',
			objects: [
				stream(
					`/Type /XObject /Subtype /Image /Width ${width} /Height ${height} /ColorSpace /DeviceGray /BitsPerComponent 1 /Filter /JBIG2Decode /DecodeParms << /JBIG2Globals 7 0 R >>`,
					bytes(fixture.data),
					true,
				),
				stream('', bytes(fixture.globals!), true),
			],
		});
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

	it('reports a JBIG2 picture whose data is damaged', async () => {
		const pdf = simplePdf('q 200 0 0 100 72 500 cm /Im Do Q', {
			resources: '/XObject << /Im 6 0 R >>',
			objects: [
				stream(
					'/Type /XObject /Subtype /Image /Width 8 /Height 8 /BitsPerComponent 1 /Filter /JBIG2Decode',
					'damaged',
					true,
				),
			],
		});
		const { report } = await importPdf(pdf);
		expect(
			report.pages[0]!.notes.some((note) =>
				note.startsWith('A picture was left out: its JBIG2 data could not be decoded'),
			),
		).toBe(true);
	});
});
