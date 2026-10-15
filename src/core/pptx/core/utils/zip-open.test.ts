import { describe, expect, it } from 'vitest';
import { PptxHandler } from '../../index';
import { deflateZeros, lyingZipOf } from '../../../opc/__fixtures__/lying-zip';
import { ZipBombError } from '../core/types';
import { openNestedZip } from './zip-open';

describe('pptx zip guards', () => {
	it('rejects a deck whose part lies about its size before inflating it', async () => {
		const bytes = lyingZipOf([
			{ name: 'ppt/presentation.xml', data: await deflateZeros(64 << 20), declared: 100 },
		]);
		const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
		const handler = new PptxHandler();
		await expect(handler.load(buffer as ArrayBuffer)).rejects.toThrow(/inflates past/);
	});

	it('rejects a nested package whose entry lies about its size', async () => {
		const bytes = lyingZipOf([
			{ name: 'xl/workbook.xml', data: await deflateZeros(64 << 20), declared: 100 },
		]);
		const zip = await openNestedZip(bytes);
		await expect(zip.file('xl/workbook.xml')!.async('string')).rejects.toBeInstanceOf(ZipBombError);
	});
});
