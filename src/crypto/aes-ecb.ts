/**
 * AES-ECB over Web Crypto, which only offers CBC: the ECMA-376 Standard scheme ([MS-OFFCRYPTO]
 * 2.3.4.7-2.3.4.9) encrypts the password verifier and the package with AES in ECB mode.
 *
 * Decryption runs one CBC pass with a zero IV and then undoes the chaining (block i was XORed
 * with ciphertext block i - 1). Encryption encrypts each block on its own (CBC with a zero IV over
 * one block is ECB), in parallel batches.
 */

import { aesCbcDecryptRaw, getSubtle } from '@christophervr/ole2/utils/ooxml-crypto-primitives';

const BLOCK = 16;
const BATCH = 4096;

/** AES-ECB decryption of block-aligned `data` (a trailing partial block is zero-padded). */
export async function aesEcbDecrypt(key: Uint8Array, data: Uint8Array): Promise<Uint8Array> {
	const plain = await aesCbcDecryptRaw(key, new Uint8Array(BLOCK), data);
	for (let at = BLOCK; at < plain.length && at < data.length; at++) {
		plain[at] = plain[at]! ^ data[at - BLOCK]!;
	}
	return plain;
}

/** AES-ECB encryption of `data`, zero-padded to the block size. */
export async function aesEcbEncrypt(key: Uint8Array, data: Uint8Array): Promise<Uint8Array> {
	const padded = new Uint8Array(Math.ceil(data.length / BLOCK) * BLOCK);
	padded.set(data);
	const subtle = getSubtle();
	const cryptoKey = await subtle.importKey('raw', key as BufferSource, { name: 'AES-CBC' }, false, [
		'encrypt',
	]);
	const iv = new Uint8Array(BLOCK);
	const out = new Uint8Array(padded.length);
	for (let start = 0; start < padded.length; start += BLOCK * BATCH) {
		const end = Math.min(padded.length, start + BLOCK * BATCH);
		const jobs: Promise<void>[] = [];
		for (let at = start; at < end; at += BLOCK) {
			const block = padded.slice(at, at + BLOCK);
			jobs.push(
				subtle.encrypt({ name: 'AES-CBC', iv }, cryptoKey, block).then((result) => {
					// Web Crypto appends a PKCS#7 padding block: keep the first block only.
					out.set(new Uint8Array(result, 0, BLOCK), at);
				}),
			);
		}
		await Promise.all(jobs);
	}
	return out;
}
