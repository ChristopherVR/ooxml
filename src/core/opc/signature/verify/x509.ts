/**
 * The few X.509 certificate fields signature verification needs, read from DER (ITU-T X.690) with
 * no platform crypto: the RSA public key, the issuer and subject names, the serial number and the
 * validity period. Anything this reader does not understand makes it return `undefined`.
 */

import { decodeBase64 } from '../../../digest/index';
import { SIGNATURE_LIMITS } from './limits';

interface Tlv {
	tag: number;
	start: number;
	end: number;
}

class DerError extends Error {}

function readTlv(bytes: Uint8Array, offset: number, limit = bytes.length): Tlv {
	const tag = bytes[offset];
	let lengthByte = bytes[offset + 1];
	if (tag === undefined || lengthByte === undefined || (tag & 0x1f) === 0x1f)
		throw new DerError('Truncated or high-tag-number DER.');
	let start = offset + 2;
	let length = lengthByte;
	if (lengthByte & 0x80) {
		const count = lengthByte & 0x7f;
		if (count === 0 || count > 4) throw new DerError('Unsupported DER length.');
		length = 0;
		for (let index = 0; index < count; index += 1) {
			lengthByte = bytes[start + index];
			if (lengthByte === undefined) throw new DerError('Truncated DER length.');
			length = length * 256 + lengthByte;
		}
		start += count;
	}
	const end = start + length;
	if (end > limit) throw new DerError('DER value runs past its container.');
	return { tag, start, end };
}

/** The child values of a constructed value. */
function children(bytes: Uint8Array, parent: Tlv): Tlv[] {
	const result: Tlv[] = [];
	for (let offset = parent.start; offset < parent.end;) {
		const child = readTlv(bytes, offset, parent.end);
		result.push(child);
		offset = child.end;
	}
	return result;
}

function expectTag(value: Tlv | undefined, tag: number): Tlv {
	if (!value || value.tag !== tag) throw new DerError(`Expected DER tag ${tag}.`);
	return value;
}

const SEQUENCE = 0x30;
const SET = 0x31;
const INTEGER = 0x02;
const BIT_STRING = 0x03;
const OBJECT_IDENTIFIER = 0x06;

const RSA_ENCRYPTION = '1.2.840.113549.1.1.1';

const NAME_ATTRIBUTES: Readonly<Record<string, string>> = {
	'2.5.4.3': 'CN',
	'2.5.4.4': 'SN',
	'2.5.4.5': 'SERIALNUMBER',
	'2.5.4.12': 'T',
	'2.5.4.42': 'GN',
	'2.5.4.6': 'C',
	'2.5.4.7': 'L',
	'2.5.4.8': 'ST',
	'2.5.4.10': 'O',
	'2.5.4.11': 'OU',
	'1.2.840.113549.1.9.1': 'E',
	'0.9.2342.19200300.100.1.25': 'DC',
};

function objectIdentifier(bytes: Uint8Array, value: Tlv): string {
	const parts: number[] = [];
	let current = 0;
	for (let index = value.start; index < value.end; index += 1) {
		const byte = bytes[index]!;
		current = current * 128 + (byte & 0x7f);
		if (!(byte & 0x80)) {
			if (parts.length === 0)
				parts.push(
					current < 80 ? Math.floor(current / 40) : 2,
					current < 80 ? current % 40 : current - 80,
				);
			else parts.push(current);
			current = 0;
		}
	}
	return parts.join('.');
}

function text(bytes: Uint8Array, value: Tlv): string {
	const content = bytes.subarray(value.start, value.end);
	switch (value.tag) {
		case 0x1e: // BMPString
			return new TextDecoder('utf-16be').decode(content);
		case 0x14: // TeletexString, read as Latin-1 like most implementations
			return Array.from(content, (byte) => String.fromCharCode(byte)).join('');
		default: // UTF8String, PrintableString, IA5String, VisibleString
			return new TextDecoder('utf-8').decode(content);
	}
}

/** A distinguished name as `CN=..., O=...`, in the order the certificate stores it. */
function distinguishedName(bytes: Uint8Array, name: Tlv): string {
	const parts: string[] = [];
	for (const relativeName of children(bytes, expectTag(name, SEQUENCE))) {
		for (const attribute of children(bytes, expectTag(relativeName, SET))) {
			const [type, value] = children(bytes, expectTag(attribute, SEQUENCE));
			const oid = objectIdentifier(bytes, expectTag(type, OBJECT_IDENTIFIER));
			if (value) parts.push(`${NAME_ATTRIBUTES[oid] ?? oid}=${text(bytes, value)}`);
		}
	}
	return parts.join(', ');
}

function time(bytes: Uint8Array, value: Tlv | undefined): string | undefined {
	if (!value || (value.tag !== 0x17 && value.tag !== 0x18)) return undefined;
	const raw = text(bytes, value);
	const match = /^(\d{2}|\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})Z$/u.exec(raw);
	if (!match) return undefined;
	const [, yearText, month, day, hour, minute, second] = match;
	let year = Number(yearText);
	if (yearText!.length === 2) year += year < 50 ? 2000 : 1900;
	const date = new Date(
		Date.UTC(year, Number(month) - 1, Number(day), Number(hour), Number(minute), Number(second)),
	);
	return Number.isNaN(date.getTime()) ? undefined : date.toISOString();
}

const hex = (bytes: Uint8Array): string =>
	Array.from(bytes, (byte) => byte.toString(16).padStart(2, '0')).join('');

/** The content of an INTEGER without its leading zero bytes. */
function magnitude(bytes: Uint8Array, value: Tlv): Uint8Array {
	let start = value.start;
	while (start < value.end && bytes[start] === 0) start += 1;
	return bytes.subarray(start, value.end);
}

const bitLength = (magnitude: Uint8Array): number =>
	magnitude.length === 0 ? 0 : (magnitude.length - 1) * 8 + (32 - Math.clz32(magnitude[0]!));

const toBigInt = (magnitude: Uint8Array): bigint =>
	magnitude.length === 0 ? 0n : BigInt(`0x${hex(magnitude)}`);

/** An RSA public key. */
export interface RsaPublicKey {
	modulus: bigint;
	exponent: bigint;
	/** Length of the modulus in bytes. */
	size: number;
}

/** What the verifier reads from a certificate. */
export interface CertificateFields {
	subject: string;
	issuer: string;
	serialNumber: string;
	validFrom?: string | undefined;
	validTo?: string | undefined;
	/**
	 * Set when the subject key is an RSA key within {@link SIGNATURE_LIMITS}: a modulus of 1024 to
	 * 16384 bits and an exponent of at most 2^32.
	 */
	rsaPublicKey?: RsaPublicKey | undefined;
}

function rsaKey(bytes: Uint8Array, subjectPublicKeyInfo: Tlv): RsaPublicKey | undefined {
	const [algorithm, key] = children(bytes, expectTag(subjectPublicKeyInfo, SEQUENCE));
	const [oid] = children(bytes, expectTag(algorithm, SEQUENCE));
	if (objectIdentifier(bytes, expectTag(oid, OBJECT_IDENTIFIER)) !== RSA_ENCRYPTION)
		return undefined;
	const bitString = expectTag(key, BIT_STRING);
	if (bytes[bitString.start] !== 0) return undefined;
	const sequence = readTlv(bytes, bitString.start + 1, bitString.end);
	const [modulusInteger, exponentInteger] = children(bytes, expectTag(sequence, SEQUENCE));
	const modulus = magnitude(bytes, expectTag(modulusInteger, INTEGER));
	const exponentBytes = magnitude(bytes, expectTag(exponentInteger, INTEGER));
	// Sizes are checked on the bytes, so an oversized key never becomes a number.
	const modulusBits = bitLength(modulus);
	if (
		modulusBits < SIGNATURE_LIMITS.minRsaModulusBits ||
		modulusBits > SIGNATURE_LIMITS.maxRsaModulusBits ||
		exponentBytes.length === 0 ||
		exponentBytes.length > 5
	)
		return undefined;
	const exponent = toBigInt(exponentBytes);
	if (exponent > SIGNATURE_LIMITS.maxRsaExponent) return undefined;
	return { modulus: toBigInt(modulus), exponent, size: modulus.length };
}

/** Read the fields of a DER certificate, or `undefined` when it is not one this reader parses. */
export function readCertificate(der: Uint8Array): CertificateFields | undefined {
	try {
		const certificate = expectTag(readTlv(der, 0), SEQUENCE);
		const [tbs] = children(der, certificate);
		const fields = children(der, expectTag(tbs, SEQUENCE));
		const offset = fields[0]?.tag === 0xa0 ? 1 : 0;
		const [serial, , issuer, validity, subject, spki] = fields.slice(offset);
		const [notBefore, notAfter] = children(der, expectTag(validity, SEQUENCE));
		return {
			serialNumber: hex(der.subarray(expectTag(serial, INTEGER).start, serial!.end)),
			issuer: distinguishedName(der, expectTag(issuer, SEQUENCE)),
			subject: distinguishedName(der, expectTag(subject, SEQUENCE)),
			validFrom: time(der, notBefore),
			validTo: time(der, notAfter),
			rsaPublicKey: rsaKey(der, expectTag(spki, SEQUENCE)),
		};
	} catch (error) {
		if (error instanceof DerError) return undefined;
		throw error;
	}
}

/** {@link readCertificate} of a Base64 DER certificate. */
export function readCertificateBase64(base64: string): CertificateFields | undefined {
	const der = decodeBase64(base64);
	return der ? readCertificate(der) : undefined;
}
