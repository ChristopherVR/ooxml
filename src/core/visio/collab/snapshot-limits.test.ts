import { deflateRawSync } from 'node:zlib';

import { describe, expect, it } from 'vitest';

import { deflateZeros, lyingZipOf } from '../../opc/__fixtures__/lying-zip';
import { createVsdx } from '../index';
import { readZip, snapshotDrawing } from './snapshot';

describe('reading a local drawing for a snapshot', () => {
	it('refuses a part over the limit a drawing is opened with', async () => {
		const bomb = lyingZipOf([
			{ name: '_rels/.rels', data: deflateRawSync('<a/>'), declared: 300 << 20 },
		]);
		await expect(snapshotDrawing(bomb, {} as never)).rejects.toMatchObject({
			name: 'VisioPackageError',
			code: 'LIMIT_ENTRY',
		});
	});

	it('stops at a part that inflates past its declared size', async () => {
		const bomb = lyingZipOf([
			{ name: '_rels/.rels', data: await deflateZeros(32 << 20), declared: 16 },
		]);
		await expect(snapshotDrawing(bomb, {} as never)).rejects.toMatchObject({
			code: 'ZIP_MISMATCH',
		});
	});

	it('still reads an ordinary drawing', async () => {
		const zip = await readZip(await createVsdx());
		expect(await zip.file('_rels/.rels')?.async('string')).toContain('Relationship');
	});
});
