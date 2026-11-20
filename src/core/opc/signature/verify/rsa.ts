/**
 * RSASSA-PKCS1-v1_5 signature verification (RFC 8017, section 8.2.2) in pure TypeScript. Only the
 * public operation runs here, so no secret is involved and timing reveals nothing.
 *
 * The encoded message is rebuilt from the expected digest and compared whole, never parsed, so a
 * signature with extra or misplaced bytes cannot pass.
 */

import { digestSync } from '../../../digest/index';
import type { RsaPublicKey } from './x509';

/** DER `DigestInfo` prefixes (RFC 8017, section 9.2, note 1), with the NULL parameters. */
const DIGEST_INFO_PREFIX: Readonly<Record<string, readonly number[]>> = {
	'SHA-1': [
		0x30, 0x21, 0x30, 0x09, 0x06, 0x05, 0x2b, 0x0e, 0x03, 0x02, 0x1a, 0x05, 0x00, 0x04, 0x14,
	],
	'SHA-256': [
		0x30, 0x31, 0x30, 0x0d, 0x06, 0x09, 0x60, 0x86, 0x48, 0x01, 0x65, 0x03, 0x04, 0x02, 0x01, 0x05,
		0x00, 0x04, 0x20,
	],
	'SHA-384': [
		0x30, 0x41, 0x30, 0x0d, 0x06, 0x09, 0x60, 0x86, 0x48, 0x01, 0x65, 0x03, 0x04, 0x02, 0x02, 0x05,
		0x00, 0x04, 0x30,
	],
	'SHA-512': [
		0x30, 0x51, 0x30, 0x0d, 0x06, 0x09, 0x60, 0x86, 0x48, 0x01, 0x65, 0x03, 0x04, 0x02, 0x03, 0x05,
		0x00, 0x04, 0x40,
	],
};

/** The same prefix without the NULL parameters, which RFC 8017 also lets a signer write. */
function withoutNullParameters(prefix: readonly number[]): number[] {
	const result = [...prefix];
	result.splice(result.length - 4, 2);
	result[1]! -= 2;
	result[3]! -= 2;
	return result;
}

function modPow(base: bigint, exponent: bigint, modulus: bigint): bigint {
	let result = 1n;
	let factor = base % modulus;
	for (let rest = exponent; rest > 0n; rest >>= 1n) {
		if (rest & 1n) result = (result * factor) % modulus;
		factor = (factor * factor) % modulus;
	}
	return result;
}

function toBytes(value: bigint, size: number): Uint8Array | undefined {
	const hex = value.toString(16).padStart(size * 2, '0');
	if (hex.length > size * 2) return undefined;
	const bytes = new Uint8Array(size);
	for (let index = 0; index < size; index += 1)
		bytes[index] = Number.parseInt(hex.slice(index * 2, index * 2 + 2), 16);
	return bytes;
}

function encodedMessage(digestInfo: readonly number[], size: number): Uint8Array | undefined {
	const padding = size - digestInfo.length - 3;
	if (padding < 8) return undefined;
	return Uint8Array.from([
		0x00,
		0x01,
		...new Array<number>(padding).fill(0xff),
		0x00,
		...digestInfo,
	]);
}

const equal = (left: Uint8Array, right: Uint8Array): boolean =>
	left.length === right.length && left.every((byte, index) => byte === right[index]);

/**
 * Whether `signature` is a valid RSASSA-PKCS1-v1_5 signature of `message` under `key` with the
 * digest `hash` (`SHA-1`, `SHA-256`, `SHA-384` or `SHA-512`).
 */
export function verifyRsaPkcs1(
	key: RsaPublicKey,
	hash: string,
	message: Uint8Array,
	signature: Uint8Array,
): boolean {
	const prefix = DIGEST_INFO_PREFIX[hash];
	const digest = digestSync(hash, message);
	if (!prefix || !digest || signature.length !== key.size) return false;
	let value = 0n;
	for (const byte of signature) value = (value << 8n) | BigInt(byte);
	if (value >= key.modulus) return false;
	const recovered = toBytes(modPow(value, key.exponent, key.modulus), key.size);
	if (!recovered) return false;
	return [prefix, withoutNullParameters(prefix)].some((candidate) => {
		const expected = encodedMessage([...candidate, ...digest], key.size);
		return expected !== undefined && equal(recovered, expected);
	});
}
