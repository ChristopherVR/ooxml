import { readFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { describe, expect, it } from 'vitest';
import { buildOle2 } from '@christophervr/ole2/ole2-parser-write';
import { loadVisio } from './load.js';
import { parseLegacyVsd } from './legacy.js';

const fixture = new URL('./__fixtures__/legacy-vsd/owned-v11.vsd', import.meta.url);

describe('authored VSD v11 shared decoder integration', () => {
	it('decodes actual compressed compound-file bytes into an honest preview scene', async () => {
		const bytes = new Uint8Array(await readFile(fixture));
		expect(createHash('sha256').update(bytes).digest('hex')).toBe(
			'a3781bcdaadfb48f3e9669808930013197b541af164dd9e44932da6d744ef36e',
		);
		const original = bytes.slice();
		const scene = await loadVisio(bytes);
		expect(scene.format).toBe('vsd');
		expect(scene.pages).toHaveLength(1);
		const page = scene.pages[0]!;
		expect(page).toMatchObject({ width: 8, height: 11, isBackground: false });
		expect(page.shapes).toHaveLength(1);
		const shape = page.shapes[0]!;
		expect(shape).toMatchObject({ id: '7', width: 4, height: 2 });
		expect(shape.geometry).toEqual([{ path: 'M 0 0 L 4 2', fill: false, stroke: true }]);
		expect(shape.transform).toEqual([1, 0, -0, 1, 0, 2]);
		expect(shape.text.plainText).toBe('Hello\n');
		expect(scene.diagnostics.map(({ code }) => code)).toEqual([
			'legacy-vsd-stored-values',
			'legacy-vsd-style-fallback',
			'legacy-vsd-text-fallback',
		]);
		expect(bytes).toEqual(original);
	});
	it('honors a byte view with a nonzero offset', async () => {
		const bytes = new Uint8Array(await readFile(fixture));
		const padded = new Uint8Array(bytes.length + 12);
		padded.set(bytes, 5);
		expect(parseLegacyVsd(padded.subarray(5, 5 + bytes.length))).toEqual(parseLegacyVsd(bytes));
	});
	it('rejects a different compound-file format instead of inventing a scene', async () => {
		const bytes = new Uint8Array(buildOle2(new Map([['Workbook', Uint8Array.of(1, 2, 3)]])));
		await expect(loadVisio(bytes)).rejects.toThrow(/Cannot decode legacy VSD/);
	});
});
