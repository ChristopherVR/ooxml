// PowerPoint-side tests of the OOXML package encryption, which moved to the shared crypto area
// (src/crypto, tested in src/crypto/crypto.test.ts): the PowerPoint-authored fixture, the
// compound-file container and the compatibility names this area keeps (decryptPptx, ...).
import { readFile } from 'node:fs/promises';

import { describe, it, expect } from 'vitest';

import * as crypto from '../../../crypto/index.js';
import { EncryptedFileError, detectFileFormat } from './encryption-detection';
import { parseOle2, buildOle2, Ole2ParseError } from './ole2-parser';
import { decryptPptx, encryptPptx, verifyPassword, IncorrectPasswordError } from './ooxml-crypto';
import { encryptPptxStandard } from './ooxml-crypto-encrypt';
import { IncorrectPasswordError as ErrorsModuleIncorrectPasswordError } from './ooxml-crypto-errors';

const TEST_SPIN_COUNT = 100;

const zip = (): ArrayBuffer => {
	const bytes = new Uint8Array(5000);
	bytes.set([0x50, 0x4b, 0x03, 0x04]);
	for (let i = 4; i < bytes.length; i++) bytes[i] = i % 251;
	return bytes.buffer;
};

describe('encryptPptx / decryptPptx (PowerPoint fixture)', () => {
	it('decrypts the password-protected PowerPoint fixture', async () => {
		const fixture = await readFile(
			new URL(
				'../../__tests__/fixtures/e2e/Password_Protected_123_8_Slides_2_3_MB_927e34cd0c.pptx',
				import.meta.url,
			),
		);
		const encrypted = fixture.buffer.slice(
			fixture.byteOffset,
			fixture.byteOffset + fixture.byteLength,
		) as ArrayBuffer;

		const decrypted = await decryptPptx(encrypted, '123');

		expect(new Uint8Array(decrypted).subarray(0, 4)).toStrictEqual(
			new Uint8Array([0x50, 0x4b, 0x03, 0x04]),
		);
		// The other crypto cases weaken the KDF, but this one cannot: the spin
		// count is baked into the PowerPoint-authored fixture. Full-strength
		// derivation over a 2.3 MB package runs ~13s alone and longer when the
		// suite saturates the machine, so the budget is generous on purpose.
	}, 120_000);
});

describe('compatibility names', () => {
	it('keeps the error classes of the shared area', () => {
		expect(IncorrectPasswordError).toBe(crypto.IncorrectPasswordError);
		expect(ErrorsModuleIncorrectPasswordError).toBe(crypto.IncorrectPasswordError);
		expect(EncryptedFileError).toBe(crypto.EncryptedFileError);
		expect(detectFileFormat).toBe(crypto.detectFileFormat);
	});

	it('round-trips ArrayBuffers through the agile and Standard schemes', async () => {
		const plain = zip();
		const agile = await encryptPptx(plain, 'pw', { spinCount: TEST_SPIN_COUNT });
		expect(agile).toBeInstanceOf(ArrayBuffer);
		expect(await verifyPassword(agile, 'pw')).toBe(true);
		expect(new Uint8Array(await decryptPptx(agile, 'pw'))).toStrictEqual(new Uint8Array(plain));
		const standard = await encryptPptxStandard(plain, 'pw');
		const decrypted = await crypto.decryptOoxmlPackage(standard, 'pw');
		expect(decrypted).toStrictEqual(new Uint8Array(plain));
		await expect(decryptPptx(standard, 'nope')).rejects.toBeInstanceOf(IncorrectPasswordError);
	});
});

describe('oLE2 container round-trip', () => {
	it('round-trips a simple stream through buildOle2 and parseOle2', () => {
		const testData = new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8]);
		const streams = new Map<string, Uint8Array>();
		streams.set('TestStream', testData);

		const ole2Buf = buildOle2(streams);
		const parsed = parseOle2(ole2Buf);
		const extracted = parsed.getStream('TestStream');

		expect(extracted).toBeDefined();
		expect(extracted!).toHaveLength(testData.length);
		expect(new Uint8Array(extracted!)).toStrictEqual(testData);
	});

	it('round-trips multiple streams', () => {
		const stream1 = new Uint8Array([0xaa, 0xbb, 0xcc]);
		const stream2 = new Uint8Array([0xdd, 0xee, 0xff]);
		const streams = new Map<string, Uint8Array>();
		streams.set('First', stream1);
		streams.set('Second', stream2);

		const ole2Buf = buildOle2(streams);
		const parsed = parseOle2(ole2Buf);

		const extracted1 = parsed.getStream('First');
		const extracted2 = parsed.getStream('Second');

		expect(extracted1).toBeDefined();
		expect(extracted2).toBeDefined();
		expect(new Uint8Array(extracted1!)).toStrictEqual(stream1);
		expect(new Uint8Array(extracted2!)).toStrictEqual(stream2);
	});

	it('round-trips a large stream (> 4096 bytes, mini stream cutoff)', () => {
		const largeData = new Uint8Array(8192);
		for (let i = 0; i < largeData.length; i++) {
			largeData[i] = i % 256;
		}
		const streams = new Map<string, Uint8Array>();
		streams.set('LargeStream', largeData);

		const ole2Buf = buildOle2(streams);
		const parsed = parseOle2(ole2Buf);
		const extracted = parsed.getStream('LargeStream');

		expect(extracted).toBeDefined();
		expect(extracted!).toHaveLength(largeData.length);
		expect(new Uint8Array(extracted!)).toStrictEqual(largeData);
	});

	it('rejects non-OLE2 data with Ole2ParseError', () => {
		const notOle2 = new ArrayBuffer(100);
		expect(() => parseOle2(notOle2)).toThrow(Ole2ParseError);
	});
});
