/**
 * EncryptionInfo stream parsing ([MS-OFFCRYPTO] 2.3.4.5 Standard, 2.3.4.10 agile).
 *
 * Moved from `src/core/pptx/core/utils/ooxml-crypto-key-derivation.ts` (see PROVENANCE.md).
 */

import { base64Decode } from '@christophervr/ole2/utils/ooxml-crypto-primitives';
import { parseXml, type XmlElement } from '../xml/xml';
import type { EncryptionInfo, StandardEncryptionInfo } from './types';
import { assertValidAgileEncryptionInfo } from './validate';

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

const PASSWORD_KEY_ENCRYPTOR_URI = 'http://schemas.microsoft.com/office/2006/keyEncryptor/password';

/** The element children of `parent` named `localName` (any prefix). */
function childrenNamed(parent: XmlElement, localName: string): XmlElement[] {
	const found: XmlElement[] = [];
	for (let node = parent.firstChild; node; node = node.nextSibling) {
		if (node.nodeType === 1 && (node as XmlElement).localName === localName) {
			found.push(node as XmlElement);
		}
	}
	return found;
}

/** The single child named `localName`, or undefined; more than one is ambiguous and rejected. */
function singleChild(
	parent: XmlElement,
	localName: string,
	mustExist: boolean,
): XmlElement | undefined {
	const matches = childrenNamed(parent, localName);
	if (matches.length > 1) throw new Error(`The EncryptionInfo has several ${localName} elements.`);
	if (matches.length === 0 && mustExist) {
		throw new Error(`The EncryptionInfo has no ${localName} element.`);
	}
	return matches[0];
}

/** A required, non-empty attribute. */
function required(element: XmlElement, attr: string): string {
	const value = element.getAttribute(attr);
	if (!value) throw new Error(`The EncryptionInfo ${element.localName} has no ${attr}.`);
	return value;
}

/** A required attribute that must be a plain decimal integer (NaN becomes an error). */
function requiredInt(element: XmlElement, attr: string): number {
	const value = required(element, attr);
	if (!/^\d{1,10}$/.test(value)) {
		throw new Error(`The EncryptionInfo ${element.localName} ${attr} is not a number: ${value}.`);
	}
	return Number(value);
}

function parseKeyBlock(element: XmlElement) {
	return {
		saltSize: requiredInt(element, 'saltSize'),
		blockSize: requiredInt(element, 'blockSize'),
		keyBits: requiredInt(element, 'keyBits'),
		hashSize: requiredInt(element, 'hashSize'),
		cipherAlgorithm: required(element, 'cipherAlgorithm'),
		cipherChaining: required(element, 'cipherChaining'),
		hashAlgorithm: required(element, 'hashAlgorithm'),
		saltValue: base64Decode(required(element, 'saltValue')),
	};
}

/**
 * Parse agile encryption info (Office 2010+ XML-based format). The XML is parsed with the shared
 * `xml` area; `keyData` must be unique, the password key encryptor is selected by its `uri`
 * (never by position), and the cipher, chaining and hash names are checked against what the
 * scheme allows.
 */
function parseAgileEncryptionInfo(data: Uint8Array): EncryptionInfo {
	// Skip version (4 bytes) and reserved (4 bytes).
	const root = parseXml(new TextDecoder('utf-8').decode(data.subarray(8)), {
		label: 'EncryptionInfo',
	}).documentElement;
	if (root.localName !== 'encryption') {
		throw new Error('The EncryptionInfo root element is not <encryption>.');
	}

	const keyData = parseKeyBlock(singleChild(root, 'keyData', true)!);
	const integrity = singleChild(root, 'dataIntegrity', false);

	const encryptors = childrenNamed(singleChild(root, 'keyEncryptors', true)!, 'keyEncryptor');
	const passwordEncryptors = encryptors.filter(
		(encryptor) => encryptor.getAttribute('uri') === PASSWORD_KEY_ENCRYPTOR_URI,
	);
	if (passwordEncryptors.length !== 1) {
		throw new Error(
			passwordEncryptors.length === 0
				? 'The EncryptionInfo has no password key encryptor.'
				: 'The EncryptionInfo has several password key encryptors.',
		);
	}
	const pke = singleChild(passwordEncryptors[0]!, 'encryptedKey', true)!;

	const info: EncryptionInfo = {
		version: { major: 4, minor: 4 },
		isAgile: true,
		keyData,
		...(integrity
			? {
					dataIntegrity: {
						encryptedHmacKey: base64Decode(required(integrity, 'encryptedHmacKey')),
						encryptedHmacValue: base64Decode(required(integrity, 'encryptedHmacValue')),
					},
				}
			: {}),
		passwordKeyEncryptor: {
			...parseKeyBlock(pke),
			spinCount: requiredInt(pke, 'spinCount'),
			encryptedVerifierHashInput: base64Decode(required(pke, 'encryptedVerifierHashInput')),
			encryptedVerifierHashValue: base64Decode(required(pke, 'encryptedVerifierHashValue')),
			encryptedKeyValue: base64Decode(required(pke, 'encryptedKeyValue')),
		},
	};
	assertValidAgileEncryptionInfo(info);
	return info;
}
