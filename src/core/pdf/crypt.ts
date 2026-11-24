/**
 * The PDF standard security handler (ISO 32000-1 section 7.6, ISO 32000-2 for AES-256): verifies
 * a password as the user or owner password and decrypts strings and streams with RC4 (40 to 128
 * bit), AES-128 or AES-256. An empty user password (a file that only restricts what a reader may
 * do) opens without asking.
 */
import {
	aesCbcDecryptRaw,
	aesCbcEncryptNoPad,
	getSubtle,
} from '@christophervr/ole2/utils/ooxml-crypto-primitives';
import { md5 } from '../digest/md5';
import { sha256 } from '../digest/sha256';
import { sha384, sha512 } from '../digest/sha512';
import {
	isDict,
	isName,
	isNumber,
	PdfName,
	PdfString,
	type PdfDict,
	type PdfValue,
} from './objects';
import { PdfPasswordError, type PdfPermissions } from './crypt-types';

export { PdfPasswordError, type PdfPermissions } from './crypt-types';

const PAD = Uint8Array.from([
	0x28, 0xbf, 0x4e, 0x5e, 0x4e, 0x75, 0x8a, 0x41, 0x64, 0x00, 0x4e, 0x56, 0xff, 0xfa, 0x01, 0x08,
	0x2e, 0x2e, 0x00, 0xb6, 0xd0, 0x68, 0x3e, 0x80, 0x2f, 0x0c, 0xa9, 0xfe, 0x64, 0x53, 0x69, 0x7a,
]);

const concat = (...parts: Uint8Array[]): Uint8Array => {
	const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0));
	let at = 0;
	for (const part of parts) {
		out.set(part, at);
		at += part.length;
	}
	return out;
};

export function rc4(key: Uint8Array, data: Uint8Array): Uint8Array {
	const s = Uint8Array.from({ length: 256 }, (_, index) => index);
	for (let i = 0, j = 0; i < 256; i++) {
		j = (j + s[i]! + key[i % key.length]!) & 0xff;
		[s[i], s[j]] = [s[j]!, s[i]!];
	}
	const out = new Uint8Array(data.length);
	for (let at = 0, i = 0, j = 0; at < data.length; at++) {
		i = (i + 1) & 0xff;
		j = (j + s[i]!) & 0xff;
		[s[i], s[j]] = [s[j]!, s[i]!];
		out[at] = data[at]! ^ s[(s[i]! + s[j]!) & 0xff]!;
	}
	return out;
}

const padded = (password: Uint8Array) => concat(password.subarray(0, 32), PAD).subarray(0, 32);
const bytesOf = (value: PdfValue | undefined) =>
	value instanceof PdfString ? value.bytes : new Uint8Array();
const int32le = (value: number) => Uint8Array.from([value, value >> 8, value >> 16, value >> 24]);

type Method = 'rc4' | 'aes' | 'identity';

/** Decrypts the objects of one encrypted file. */
export class PdfDecryptor {
	private constructor(
		readonly key: Uint8Array,
		readonly streams: Method,
		readonly strings: Method,
		readonly revision: number,
		readonly permissions: PdfPermissions,
		readonly encryptMetadata: boolean,
	) {}

	static async open(encrypt: PdfDict, id: Uint8Array, password = ''): Promise<PdfDecryptor> {
		if (!isName(encrypt.get('Filter'), 'Standard'))
			throw new Error('This PDF uses a security handler other than the standard one.');
		const v = isNumber(encrypt.get('V')) ? (encrypt.get('V') as number) : 0;
		const r = isNumber(encrypt.get('R')) ? (encrypt.get('R') as number) : 2;
		const p = isNumber(encrypt.get('P')) ? (encrypt.get('P') as number) : -1;
		const o = bytesOf(encrypt.get('O'));
		const u = bytesOf(encrypt.get('U'));
		const encryptMetadata = encrypt.get('EncryptMetadata') !== false;
		const length = isNumber(encrypt.get('Length')) ? (encrypt.get('Length') as number) : 40;
		const method = (name: string | undefined): Method => {
			if (v < 4) return 'rc4';
			if (name === 'Identity') return 'identity';
			const filter = encrypt.get('CF');
			const cf = isDict(filter) ? filter.get(name ?? '') : undefined;
			const cfm = isDict(cf) ? cf.get('CFM') : undefined;
			return isName(cfm, 'V2') ? 'rc4' : isName(cfm, 'None') ? 'identity' : 'aes';
		};
		const streamName = encrypt.get('StmF');
		const stringName = encrypt.get('StrF');
		const streams = method(streamName instanceof PdfName ? streamName.name : 'Identity');
		const strings = method(stringName instanceof PdfName ? stringName.name : 'Identity');
		const permissions = decodePermissions(p);
		const text = new TextEncoder().encode(password);
		let key: Uint8Array | undefined;
		if (r >= 5) key = await aes256Key(text.subarray(0, 127), o, u, encrypt, r);
		else {
			const n = r === 2 ? 5 : Math.min(16, Math.max(5, length / 8));
			const latin = Uint8Array.from([...password].map((char) => char.charCodeAt(0) & 0xff));
			key = legacyUserKey(latin, o, u, p, id, r, n, encryptMetadata);
			if (!key) {
				// The password may be the owner password: it unlocks the user password.
				const user = ownerToUser(latin, o, r, n);
				key = legacyUserKey(user, o, u, p, id, r, n, encryptMetadata);
			}
		}
		if (!key)
			throw new PdfPasswordError(
				password
					? 'The password is incorrect.'
					: 'This PDF is password protected. Enter the password to open it.',
				password ? 'incorrect-password' : 'password-required',
			);
		return new PdfDecryptor(key, streams, strings, r, permissions, encryptMetadata);
	}

	#objectKey(num: number, gen: number, method: Method): Uint8Array {
		if (this.revision >= 5) return this.key;
		const salt = method === 'aes' ? Uint8Array.from([0x73, 0x41, 0x6c, 0x54]) : new Uint8Array();
		const extra = Uint8Array.from([num, num >> 8, num >> 16, gen, gen >> 8]);
		return md5(concat(this.key, extra, salt)).subarray(0, Math.min(16, this.key.length + 5));
	}

	async #decrypt(data: Uint8Array, num: number, gen: number, method: Method): Promise<Uint8Array> {
		if (method === 'identity' || !data.length) return data;
		const key = this.#objectKey(num, gen, method);
		if (method === 'rc4') return rc4(key, data);
		if (data.length < 32 || data.length % 16) return new Uint8Array();
		try {
			const subtle = getSubtle();
			const cryptoKey = await subtle.importKey(
				'raw',
				key as BufferSource,
				{ name: 'AES-CBC' },
				false,
				['decrypt'],
			);
			return new Uint8Array(
				await subtle.decrypt(
					{ name: 'AES-CBC', iv: data.subarray(0, 16) as BufferSource },
					cryptoKey,
					data.subarray(16) as BufferSource,
				),
			);
		} catch {
			// Bad padding: keep the raw plaintext, as viewers do.
			return aesCbcDecryptRaw(key, data.subarray(0, 16), data.subarray(16));
		}
	}

	decryptStream(data: Uint8Array, num: number, gen: number): Promise<Uint8Array> {
		return this.#decrypt(data, num, gen, this.streams);
	}

	decryptString(data: Uint8Array, num: number, gen: number): Promise<Uint8Array> {
		return this.#decrypt(data, num, gen, this.strings);
	}
}

function decodePermissions(p: number): PdfPermissions {
	const bit = (n: number) => (p & (1 << (n - 1))) !== 0;
	return {
		print: bit(3),
		modify: bit(4),
		copy: bit(5),
		annotate: bit(6),
		fillForms: bit(9),
		extractForAccessibility: bit(10),
		assemble: bit(11),
		printHighQuality: bit(12),
	};
}

/** Algorithm 2 (the file key from a user password) checked with algorithms 4 and 5. */
function legacyUserKey(
	password: Uint8Array,
	o: Uint8Array,
	u: Uint8Array,
	p: number,
	id: Uint8Array,
	r: number,
	n: number,
	encryptMetadata: boolean,
): Uint8Array | undefined {
	let hash = md5(
		concat(
			padded(password),
			o.subarray(0, 32),
			int32le(p),
			id,
			r >= 4 && !encryptMetadata ? Uint8Array.from([255, 255, 255, 255]) : new Uint8Array(),
		),
	);
	if (r >= 3) for (let round = 0; round < 50; round++) hash = md5(hash.subarray(0, n));
	const key = hash.subarray(0, n);
	if (r === 2) return equal(rc4(key, PAD), u.subarray(0, 32)) ? key : undefined;
	let check = md5(concat(PAD, id));
	for (let round = 0; round < 20; round++)
		check = rc4(
			key.map((byte) => byte ^ round),
			check,
		);
	return equal(check.subarray(0, 16), u.subarray(0, 16)) ? key : undefined;
}

/** Algorithm 7: the user password an owner password unlocks. */
function ownerToUser(password: Uint8Array, o: Uint8Array, r: number, n: number): Uint8Array {
	let hash = md5(padded(password));
	if (r >= 3) for (let round = 0; round < 50; round++) hash = md5(hash);
	const key = hash.subarray(0, n);
	let user = o.subarray(0, 32);
	if (r === 2) return rc4(key, user);
	for (let round = 19; round >= 0; round--)
		user = rc4(
			key.map((byte) => byte ^ round),
			user,
		);
	return user;
}

const equal = (a: Uint8Array, b: Uint8Array) =>
	a.length === b.length && a.every((byte, at) => byte === b[at]);

/** Algorithm 2.B (revision 6) or SHA-256 (revision 5). */
async function hash6(
	password: Uint8Array,
	salt: Uint8Array,
	udata: Uint8Array,
	r: number,
): Promise<Uint8Array> {
	let k: Uint8Array = sha256(concat(password, salt, udata));
	if (r === 5) return k;
	for (let round = 0; ; round++) {
		const block = concat(password, k, udata);
		const k1 = new Uint8Array(block.length * 64);
		for (let copy = 0; copy < 64; copy++) k1.set(block, copy * block.length);
		const e = await aesCbcEncryptNoPad(k.subarray(0, 16), k.subarray(16, 32), k1);
		let sum = 0;
		for (let at = 0; at < 16; at++) sum += e[at]!;
		const mod = sum % 3;
		k = mod === 0 ? sha256(e) : mod === 1 ? sha384(e) : sha512(e);
		if (round >= 63 && e[e.length - 1]! <= round - 32) break;
	}
	return k.subarray(0, 32);
}

async function aes256Key(
	password: Uint8Array,
	o: Uint8Array,
	u: Uint8Array,
	encrypt: PdfDict,
	r: number,
): Promise<Uint8Array | undefined> {
	const ue = bytesOf(encrypt.get('UE'));
	const oe = bytesOf(encrypt.get('OE'));
	const zero = new Uint8Array(16);
	if (equal(await hash6(password, u.subarray(32, 40), new Uint8Array(), r), u.subarray(0, 32))) {
		const key = await hash6(password, u.subarray(40, 48), new Uint8Array(), r);
		return aesCbcDecryptRaw(key, zero, ue.subarray(0, 32));
	}
	const udata = u.subarray(0, 48);
	if (equal(await hash6(password, o.subarray(32, 40), udata, r), o.subarray(0, 32))) {
		const key = await hash6(password, o.subarray(40, 48), udata, r);
		return aesCbcDecryptRaw(key, zero, oe.subarray(0, 32));
	}
	return undefined;
}
