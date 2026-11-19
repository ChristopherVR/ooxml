import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { deflateZeros, lyingZipOf } from './__fixtures__/lying-zip';
import { diffOpcPackages } from './diff-packages';
import { ZipLimitError } from './zip-limits';

async function pack(parts: Record<string, string>, compression: 'STORE' | 'DEFLATE' = 'DEFLATE') {
	const zip = new JSZip();
	for (const [name, text] of Object.entries(parts)) zip.file(name, text);
	return zip.generateAsync({ type: 'uint8array', compression });
}

describe('diffOpcPackages', () => {
	it('lists added, removed and changed parts', async () => {
		const base = await pack({ 'a.xml': '<a/>', 'b.xml': '<b/>', 'c.xml': '<c/>' });
		const other = await pack({ 'a.xml': '<a/>', 'b.xml': '<b x="1"/>', 'd.xml': '<d/>' });
		expect(await diffOpcPackages(base, other)).toEqual({
			added: ['d.xml'],
			removed: ['c.xml'],
			changed: ['b.xml'],
		});
	});

	it('ignores compression and part order', async () => {
		const base = await pack({ 'a.xml': '<a/>', 'b.xml': '<b/>' }, 'DEFLATE');
		const other = await pack({ 'b.xml': '<b/>', 'a.xml': '<a/>' }, 'STORE');
		expect(await diffOpcPackages(base, other)).toEqual({ added: [], removed: [], changed: [] });
	});

	it('refuses a package over the zip limits before inflating it', async () => {
		const bomb = lyingZipOf([
			{ name: 'a.xml', data: await deflateZeros(1 << 20), declared: 400 << 20 },
		]);
		await expect(diffOpcPackages(bomb, bomb)).rejects.toBeInstanceOf(ZipLimitError);
	});

	it('accepts the limits a caller passes', async () => {
		const small = await pack({ 'a.xml': '<a/>' }, 'STORE');
		await expect(diffOpcPackages(small, small, { maxInputBytes: 8 })).rejects.toBeInstanceOf(
			ZipLimitError,
		);
		await expect(diffOpcPackages(small, small, { maxEntries: 5 })).resolves.toMatchObject({
			changed: [],
		});
	});

	it('stops at a part that inflates past its declared size', async () => {
		const bomb = lyingZipOf([{ name: 'a.xml', data: await deflateZeros(8 << 20), declared: 8 }]);
		await expect(diffOpcPackages(bomb, bomb)).rejects.toThrow(/inflates past/);
	});
});
