/**
 * EncryptionInfo stream parsing ([MS-OFFCRYPTO] 2.3.4.5 Standard, 2.3.4.10 agile).
 *
 * Moved from `src/pptx/core/utils/ooxml-crypto-key-derivation.ts` (see PROVENANCE.md).
 */

import { base64Decode } from '@christophervr/ole2/utils/ooxml-crypto-primitives';
import type { EncryptionInfo, StandardEncryptionInfo } from './types.js';

/**
 * Algorithm ID for RC4 in the [MS-OFFCRYPTO] EncryptionHeader `algId` field.
 *
 * RC4 is a stream cipher, so its encrypted verifier hash is stored at exactly
 * `verifierHashSize` bytes with no block-alignment padding, unlike AES (a block cipher, padded
 * up to a 16-byte boundary).
 */
export const RC4_ALG_ID = 0x6801;

/**
 * Parse the EncryptionInfo stream of an encrypted OOXML file: agile (version 4.4) or Standard
 * (versions 2.2, 3.2 and 4.2).
 *
 * @throws Error if the encryption version is unsupported.
 */
export function parseEncryptionInfo(data: Uint8Array): EncryptionInfo | StandardEncryptionInfo {
	if (data.byteLength < 8) throw new Error('The EncryptionInfo stream is truncated.');
	const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
	const versionMajor = view.getUint16(0, true);
	const versionMinor = view.getUint16(2, true);
	if (versionMajor === 4 && versionMinor === 4) return parseAgileEncryptionInfo(data);
	if ((versionMajor === 2 || versionMajor === 3 || versionMajor === 4) && versionMinor === 2) {
		return parseStandardEncryptionInfo(data);
	}
	throw new Error(
		`Unsupported encryption version: ${versionMajor}.${versionMinor}. ` +
			'Only Standard (2.2-4.2) and Agile (4.4) encryption are supported.',
	);
}

/**
 * Parse Standard encryption info (Office 2007 format). Also used for the byte-identical
 * structure embedded in a legacy `.ppt` CryptSession10Container record ([MS-OFFCRYPTO] 2.3.5.1).
 */
function parseStandardEncryptionInfo(data: Uint8Array): StandardEncryptionInfo {
	const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
	const versionMajor = view.getUint16(0, true);
	const versionMinor = view.getUint16(2, true);
	const flags = view.getUint32(4, true);
	const headerSize = view.getUint32(8, true);

	// The EncryptionHeader starts at offset 12; sizeExtra (+4) and two reserved fields are skipped.
	const headerOffset = 12;
	const hFlags = view.getUint32(headerOffset, true);
	const algId = view.getUint32(headerOffset + 8, true);
	const algIdHash = view.getUint32(headerOffset + 12, true);
	const keySize = view.getUint32(headerOffset + 16, true);
	const providerType = view.getUint32(headerOffset + 20, true);

	// The CSP name is a UTF-16LE string after the 32 fixed header bytes.
	let cspName = '';
	const cspEnd = headerOffset + headerSize;
	for (let i = headerOffset + 32; i < cspEnd - 1; i += 2) {
		const ch = view.getUint16(i, true);
		if (ch === 0) break;
		cspName += String.fromCharCode(ch);
	}

	const verifierOffset = 12 + headerSize;
	const saltSize = view.getUint32(verifierOffset, true);
	const salt = new Uint8Array(data.buffer, data.byteOffset + verifierOffset + 4, 16);
	const encryptedVerifier = new Uint8Array(data.buffer, data.byteOffset + verifierOffset + 20, 16);
	const verifierHashSize = view.getUint32(verifierOffset + 36, true);
	// RC4 (a stream cipher) stores exactly verifierHashSize bytes; AES pads to 16 bytes.
	const encryptedVerifierHashLength =
		algId === RC4_ALG_ID ? verifierHashSize : Math.ceil(verifierHashSize / 16) * 16;
	const encryptedVerifierHash = new Uint8Array(
		data.buffer,
		data.byteOffset + verifierOffset + 40,
		encryptedVerifierHashLength,
	);

	return {
		version: { major: versionMajor, minor: versionMinor },
		isAgile: false,
		isStandard: true,
		flags,
		headerSize,
		header: { flags: hFlags, algId, algIdHash, keySize, providerType, cspName },
		verifier: {
			saltSize,
			salt: new Uint8Array(salt),
			encryptedVerifier: new Uint8Array(encryptedVerifier),
			verifierHashSize,
			encryptedVerifierHash: new Uint8Array(encryptedVerifierHash),
		},
	};
}

/** The value of `attr` in the first start tag whose name contains `tag`, or ''. */
function tagAttribute(xml: string, tag: string, attr: string): string {
	const tagMatch = xml.match(new RegExp(`<[^>]*${tag}[^>]*>`, 'i'));
	if (!tagMatch) return '';
	const attrMatch = tagMatch[0].match(new RegExp(`\\b${attr}="([^"]*)"`, 'i'));
	return attrMatch?.[1] ?? '';
}

/** Parse agile encryption info (Office 2010+ XML-based format). */
function parseAgileEncryptionInfo(data: Uint8Array): EncryptionInfo {
	// Skip version (4 bytes) and reserved (4 bytes).
	const xml = new TextDecoder('utf-8').decode(data.subarray(8));
	const keyData = (attr: string): string => tagAttribute(xml, 'keyData', attr);
	const integrity = (attr: string): string => tagAttribute(xml, 'dataIntegrity', attr);
	// The password key encryptor is the (usually `p:`-prefixed) encryptedKey element.
	const pke = (attr: string): string => tagAttribute(xml, 'encryptedKey', attr);
	const int = (value: string): number => parseInt(value, 10);

	const encryptedHmacKey = integrity('encryptedHmacKey');
	return {
		version: { major: 4, minor: 4 },
		isAgile: true,
		keyData: {
			saltSize: int(keyData('saltSize')),
			blockSize: int(keyData('blockSize')),
			keyBits: int(keyData('keyBits')),
			hashSize: int(keyData('hashSize')),
			cipherAlgorithm: keyData('cipherAlgorithm'),
			cipherChaining: keyData('cipherChaining'),
			hashAlgorithm: keyData('hashAlgorithm'),
			saltValue: base64Decode(keyData('saltValue')),
		},
		...(encryptedHmacKey
			? {
					dataIntegrity: {
						encryptedHmacKey: base64Decode(encryptedHmacKey),
						encryptedHmacValue: base64Decode(integrity('encryptedHmacValue')),
					},
				}
			: {}),
		passwordKeyEncryptor: {
			saltSize: int(pke('saltSize')),
			blockSize: int(pke('blockSize')),
			keyBits: int(pke('keyBits')),
			hashSize: int(pke('hashSize')),
			cipherAlgorithm: pke('cipherAlgorithm'),
			cipherChaining: pke('cipherChaining'),
			hashAlgorithm: pke('hashAlgorithm'),
			saltValue: base64Decode(pke('saltValue')),
			spinCount: int(pke('spinCount')),
			encryptedVerifierHashInput: base64Decode(pke('encryptedVerifierHashInput')),
			encryptedVerifierHashValue: base64Decode(pke('encryptedVerifierHashValue')),
			encryptedKeyValue: base64Decode(pke('encryptedKeyValue')),
		},
	};
}
