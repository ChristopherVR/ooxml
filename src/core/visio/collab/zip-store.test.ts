import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { VisioPackage } from '../package';
import { writeStoredZip } from './zip-store';

const text = (s: string) => new TextEncoder().encode(s);

describe('the stored ZIP writer', () => {
	it('writes a package JSZip and the Visio reader both open', async () => {
		const big = new Uint8Array(70_000).map((_, i) => i % 251);
		const bytes = writeStoredZip(
			[
				['[Content_Types].xml', text('<Types/>')],
				['visio/pages/pagé.xml', text('<PageContents/>')],
				['visio/media/big.bin', big],
				['empty.txt', new Uint8Array(0)],
			],
			new WeakMap(),
		);
		const zip = await JSZip.loadAsync(bytes, { checkCRC32: true });
		expect(Object.keys(zip.files)).toEqual([
			'[Content_Types].xml',
			'visio/pages/pagé.xml',
			'visio/media/big.bin',
			'empty.txt',
		]);
		expect(await zip.file('visio/media/big.bin')!.async('uint8array')).toEqual(big);
		expect(await zip.file('visio/pages/pagé.xml')!.async('string')).toBe('<PageContents/>');
		const pkg = await VisioPackage.open(bytes);
		expect(pkg.paths()).toHaveLength(4);
		expect(await pkg.readBytes('visio/media/big.bin')).toEqual(big);
	});

	it('checksums a byte array once and writes equal bytes for equal input', () => {
		const crcs = new WeakMap<Uint8Array, number>();
		const part = text('<a/>');
		const files = [['a.xml', part]] as const;
		const first = writeStoredZip(files, crcs);
		expect(crcs.has(part)).toBe(true);
		crcs.set(part, crcs.get(part)!); // a cache hit changes nothing
		expect(writeStoredZip(files, crcs)).toEqual(first);
		expect(writeStoredZip(files, new WeakMap())).toEqual(first);
	});
});
