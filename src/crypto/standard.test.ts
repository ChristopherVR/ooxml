// The ECMA-376 Standard scheme checked against an independent node:crypto implementation of
// [MS-OFFCRYPTO] 2.3.4.7 (key derivation, always through X1/X2/X3) and 2.3.4.15 (AES-ECB). Both
// were wrong before the move (a truncated 128-bit key and AES-CBC), so Excel refused the files.
import { createDecipheriv, createHash } from 'node:crypto';
import { parseOle2 } from '@christophervr/ole2/ole2-parser-read';
import { describe, expect, it } from 'vitest';
import { decryptOoxmlPackage, encryptOoxmlPackage, parseEncryptionInfo } from './index.js';
import type { StandardEncryptionInfo } from './types.js';

const sha1 = (...parts: Uint8Array[]): Buffer => {
	const hash = createHash('sha1');
	for (const part of parts) hash.update(part);
	return hash.digest();
};

function referenceKey(password: string, salt: Uint8Array, keyBits: number): Buffer {
	let h = sha1(salt, Buffer.from(password, 'utf16le'));
	const counter = Buffer.alloc(4);
	for (let i = 0; i < 50000; i++) {
		counter.writeUInt32LE(i);
		h = sha1(counter, h);
	}
	const final = sha1(h, Buffer.alloc(4));
	const pad = (value: number): Buffer => {
		const buffer = Buffer.alloc(64, value);
		for (let i = 0; i < final.length; i++) buffer[i] = final[i]! ^ value;
		return sha1(buffer);
	};
	return Buffer.concat([pad(0x36), pad(0x5c)]).subarray(0, keyBits / 8);
}

const ecbDecrypt = (key: Buffer, data: Uint8Array): Buffer => {
	const decipher = createDecipheriv(`aes-${key.length * 8}-ecb`, key, null);
	decipher.setAutoPadding(false);
	return Buffer.concat([decipher.update(data), decipher.final()]);
};

describe('the Standard encryption scheme', () => {
	for (const algorithm of ['AES128', 'AES256'] as const) {
		it(`matches the reference derivation and AES-ECB (${algorithm})`, async () => {
			// Repeated 16-byte blocks: ECB encrypts equal blocks equally, CBC would not.
			const plain = new Uint8Array(100).map((_, index) => index % 16);
			const encrypted = await encryptOoxmlPackage(plain, 'open sesame', {
				encryptionScheme: 'standard',
				algorithm,
			});
			const file = parseOle2(encrypted.slice().buffer);
			const info = parseEncryptionInfo(file.getStream('EncryptionInfo')!) as StandardEncryptionInfo;
			expect(info.isStandard).toBe(true);
			const key = referenceKey('open sesame', info.verifier.salt, info.header.keySize);
			const verifier = ecbDecrypt(key, info.verifier.encryptedVerifier);
			const verifierHash = ecbDecrypt(key, info.verifier.encryptedVerifierHash);
			expect(verifierHash.subarray(0, 20)).toStrictEqual(sha1(verifier));
			const stream = file.getStream('EncryptedPackage')!;
			expect(new DataView(stream.buffer, stream.byteOffset).getUint32(0, true)).toBe(100);
			const body = stream.subarray(8);
			expect(body.subarray(0, 16)).toStrictEqual(body.subarray(16, 32));
			expect(new Uint8Array(ecbDecrypt(key, body).subarray(0, 100))).toStrictEqual(plain);
			expect(await decryptOoxmlPackage(encrypted, 'open sesame')).toStrictEqual(plain);
		}, 30_000);
	}
});
