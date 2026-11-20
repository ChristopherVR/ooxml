import { readFileSync } from 'node:fs';

import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../PptxHandler';

const FIXTURE = new URL('../../../__tests__/fixtures/embedded-assets-sample.pptx', import.meta.url);

/** `xml` as UTF-16 little-endian with a byte order mark, as ECMA-376 Part 2 allows. */
function utf16(xml: string): Uint8Array {
	const text = xml.replace(/encoding="UTF-8"/iu, 'encoding="UTF-16"');
	const bytes = new Uint8Array(2 + text.length * 2);
	bytes.set([0xff, 0xfe]);
	for (let index = 0; index < text.length; index += 1) {
		bytes[2 + index * 2] = text.charCodeAt(index) & 0xff;
		bytes[3 + index * 2] = text.charCodeAt(index) >> 8;
	}
	return bytes;
}

describe('a deck with UTF-16 parts', () => {
	it('loads its slides as it does in UTF-8', async () => {
		const original = new Uint8Array(readFileSync(FIXTURE));
		const expected = (await new PptxHandler().load(original.slice().buffer)).slides.length;
		const zip = await JSZip.loadAsync(original);
		for (const path of ['_rels/.rels', 'ppt/_rels/presentation.xml.rels'])
			zip.file(path, utf16(await zip.file(path)!.async('string')));
		const bytes = await zip.generateAsync({ type: 'arraybuffer' });
		const loaded = await new PptxHandler().load(bytes);
		expect(expected).toBeGreaterThan(0);
		expect(loaded.slides).toHaveLength(expected);
	});
});
