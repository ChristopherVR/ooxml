// Hostile EncryptionInfo and EncryptedPackage inputs: downgrade, forged sizes, spin counts and
// ambiguous XML. Each case starts from a real package encrypted with a low spin count.

import { buildOle2 } from '@christophervr/ole2/ole2-parser-write';
import { parseOle2 } from '@christophervr/ole2/ole2-parser-read';
import { describe, expect, it } from 'vitest';
import {
	DataIntegrityError,
	IncorrectPasswordError,
	decryptAgilePackage,
	decryptOoxmlPackage,
	encryptOoxmlPackage,
	inspectOoxmlEncryption,
	parseEncryptionInfo,
	verifyAgilePassword,
	verifyOoxmlPackagePassword,
	type EncryptionInfo,
} from './index';

const PASSWORD = 'correct horse';
const PLAIN = new TextEncoder().encode('PK-pretend-zip-content-'.repeat(300));

async function encrypted(): Promise<Uint8Array> {
	return encryptOoxmlPackage(PLAIN, PASSWORD, { spinCount: 100 });
}

/** Rebuild the compound file with a rewritten EncryptionInfo XML and/or package bytes. */
function rebuild(
	file: Uint8Array,
	edit: { xml?: (xml: string) => string; pkg?: (pkg: Uint8Array) => Uint8Array },
): Uint8Array {
	const ole = parseOle2(file.slice().buffer as ArrayBuffer);
	const info = ole.getStream('EncryptionInfo')!;
	const pkg = ole.getStream('EncryptedPackage')!;
	const head = info.slice(0, 8);
	const xml = new TextDecoder().decode(info.subarray(8));
	const newXml = new TextEncoder().encode(edit.xml ? edit.xml(xml) : xml);
	const newInfo = new Uint8Array(8 + newXml.length);
	newInfo.set(head);
	newInfo.set(newXml, 8);
	const streams = new Map<string, Uint8Array>();
	streams.set('EncryptionInfo', newInfo);
	streams.set('EncryptedPackage', edit.pkg ? edit.pkg(new Uint8Array(pkg)) : pkg);
	return new Uint8Array(buildOle2(streams));
}

const infoOf = (file: Uint8Array): EncryptionInfo => {
	const ole = parseOle2(file.slice().buffer as ArrayBuffer);
	return parseEncryptionInfo(ole.getStream('EncryptionInfo')!) as EncryptionInfo;
};

const stripIntegrity = (xml: string): string => xml.replace(/<dataIntegrity[^>]*\/>/u, '');

describe('missing dataIntegrity (downgrade)', () => {
	it('rejects an agile package whose dataIntegrity was stripped', async () => {
		const stripped = rebuild(await encrypted(), { xml: stripIntegrity });
		await expect(decryptOoxmlPackage(stripped, PASSWORD)).rejects.toThrow(DataIntegrityError);
	});

	it('accepts it only with allowMissingIntegrity', async () => {
		const stripped = rebuild(await encrypted(), { xml: stripIntegrity });
		const plain = await decryptOoxmlPackage(stripped, PASSWORD, { allowMissingIntegrity: true });
		expect(plain).toStrictEqual(PLAIN);
	});

	it('reports whether the ciphertext is authenticated', async () => {
		const file = await encrypted();
		expect(inspectOoxmlEncryption(file)).toStrictEqual({
			scheme: 'agile',
			integrityProtected: true,
		});
		expect(inspectOoxmlEncryption(rebuild(file, { xml: stripIntegrity })).integrityProtected).toBe(
			false,
		);
		const standard = await encryptOoxmlPackage(PLAIN, PASSWORD, { encryptionScheme: 'standard' });
		expect(inspectOoxmlEncryption(standard)).toStrictEqual({
			scheme: 'standard',
			integrityProtected: false,
		});
	});
});

describe('EncryptionInfo parameter validation', () => {
	it.each(['NaN', '0', '32', '65', '-1'])('rejects hashSize="%s"', async (value) => {
		const forged = rebuild(await encrypted(), {
			xml: (xml) => xml.replace(/hashSize="64"/gu, `hashSize="${value}"`),
		});
		await expect(decryptOoxmlPackage(forged, PASSWORD)).rejects.toThrow(/hashSize/u);
		expect(await verifyOoxmlPackagePassword(forged, PASSWORD)).toBe(false);
		expect(await verifyOoxmlPackagePassword(forged, 'wrong')).toBe(false);
	});

	it('never lets any password verify with hashSize 0 on a crafted info object', async () => {
		const info = infoOf(await encrypted());
		const forged = {
			...info,
			passwordKeyEncryptor: { ...info.passwordKeyEncryptor, hashSize: 0 },
		};
		await expect(verifyAgilePassword(forged, 'anything')).rejects.toThrow(/hashSize/u);
	});

	it.each([
		['keyBits="256"', 'keyBits="64"', /keyBits/u],
		['blockSize="16"', 'blockSize="8"', /blockSize/u],
		['saltSize="16"', 'saltSize="0"', /saltSize/u],
		['cipherAlgorithm="AES"', 'cipherAlgorithm="RC2"', /cipherAlgorithm/u],
		['cipherChaining="ChainingModeCBC"', 'cipherChaining="ChainingModeCFB"', /cipherChaining/u],
		['hashAlgorithm="SHA-512"', 'hashAlgorithm="MD5"', /hashAlgorithm/u],
	])('rejects %s -> %s', async (from, to, message) => {
		const forged = rebuild(await encrypted(), { xml: (xml) => xml.split(from).join(to) });
		await expect(decryptOoxmlPackage(forged, PASSWORD)).rejects.toThrow(message);
	});

	it('rejects a fractional or non-numeric spinCount', async () => {
		for (const value of ['1.5', 'abc', '-3']) {
			const forged = rebuild(await encrypted(), {
				xml: (xml) => xml.replace(/spinCount="\d+"/u, `spinCount="${value}"`),
			});
			await expect(decryptOoxmlPackage(forged, PASSWORD)).rejects.toThrow(/spinCount/u);
		}
	});

	it('rejects a Standard verifier hash size of zero', async () => {
		const standard = await encryptOoxmlPackage(PLAIN, PASSWORD, { encryptionScheme: 'standard' });
		const ole = parseOle2(standard.slice().buffer as ArrayBuffer);
		const info = new Uint8Array(ole.getStream('EncryptionInfo')!);
		const view = new DataView(info.buffer);
		const headerSize = view.getUint32(8, true);
		view.setUint32(12 + headerSize + 36, 0, true);
		const streams = new Map<string, Uint8Array>();
		streams.set('EncryptionInfo', info);
		streams.set('EncryptedPackage', ole.getStream('EncryptedPackage')!);
		await expect(decryptOoxmlPackage(buildOle2(streams), PASSWORD)).rejects.toThrow(
			/verifierHashSize/u,
		);
	});
});

describe('declared package size', () => {
	it('rejects a size larger than the ciphertext before allocating it', async () => {
		const file = await encrypted();
		const info = infoOf(file);
		const key = (await verifyAgilePassword(info, PASSWORD))!;
		const ole = parseOle2(file.slice().buffer as ArrayBuffer);
		const pkg = new Uint8Array(ole.getStream('EncryptedPackage')!);
		// 2^40 bytes declared (high dword 256): would be a terabyte allocation.
		new DataView(pkg.buffer).setUint32(4, 256, true);
		await expect(decryptAgilePackage(pkg, key, info)).rejects.toThrow(/declared package size/u);
	});

	it('rejects a truncated EncryptedPackage stream', async () => {
		const info = infoOf(await encrypted());
		await expect(decryptAgilePackage(new Uint8Array(4), new Uint8Array(32), info)).rejects.toThrow(
			/truncated/u,
		);
	});
});

describe('spin count limit', () => {
	it('rejects a hostile spinCount quickly and honours an explicit raise', async () => {
		const forged = rebuild(await encrypted(), {
			xml: (xml) => xml.replace(/spinCount="\d+"/u, 'spinCount="10000000"'),
		});
		const started = Date.now();
		await expect(decryptOoxmlPackage(forged, PASSWORD)).rejects.toThrow(/spin count/u);
		expect(Date.now() - started).toBeLessThan(1000);
		// Raising the limit past the spec ceiling is still refused.
		const tooBig = rebuild(await encrypted(), {
			xml: (xml) => xml.replace(/spinCount="\d+"/u, 'spinCount="10000001"'),
		});
		await expect(
			decryptOoxmlPackage(tooBig, PASSWORD, { maxSpinCount: 50_000_000 }),
		).rejects.toThrow(/spin count/u);
	});

	it('opens a package spun above the default limit when the caller raises it', async () => {
		const info = infoOf(await encrypted());
		const raised = {
			...info,
			passwordKeyEncryptor: { ...info.passwordKeyEncryptor, spinCount: 1_000_001 },
		};
		await expect(verifyAgilePassword(raised, PASSWORD)).rejects.toThrow(/spin count/u);
		// Wrong spin for the verifier: the password check fails (null), but it was allowed to run.
		expect(await verifyAgilePassword(raised, PASSWORD, { maxSpinCount: 2_000_000 })).toBeNull();
	}, 30_000);
});

describe('agile EncryptionInfo XML selection', () => {
	const DECOY =
		'<keyEncryptor uri="http://example.com/other"><p:encryptedKey spinCount="1" saltSize="16" ' +
		'blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" ' +
		'cipherChaining="ChainingModeCBC" hashAlgorithm="SHA-512" saltValue="AAAAAAAAAAAAAAAAAAAAAA==" ' +
		'encryptedVerifierHashInput="AA==" encryptedVerifierHashValue="AA==" encryptedKeyValue="AA=="/>' +
		'</keyEncryptor>';

	it('picks the password key encryptor by uri, not by position', async () => {
		const reordered = rebuild(await encrypted(), {
			xml: (xml) => xml.replace('<keyEncryptors>', `<keyEncryptors>${DECOY}`),
		});
		expect(await decryptOoxmlPackage(reordered, PASSWORD)).toStrictEqual(PLAIN);
		await expect(decryptOoxmlPackage(reordered, 'wrong')).rejects.toThrow(IncorrectPasswordError);
	});

	it('is not fooled by an attribute-bearing comment or CDATA before the real element', async () => {
		const forged = rebuild(await encrypted(), {
			xml: (xml) => xml.replace('<keyData', '<!-- <keyData saltSize="1" keyBits="8"/> --><keyData'),
		});
		expect(await decryptOoxmlPackage(forged, PASSWORD)).toStrictEqual(PLAIN);
	});

	it('rejects duplicate keyData and duplicate password encryptors', async () => {
		const dupKeyData = rebuild(await encrypted(), {
			xml: (xml) => xml.replace(/<keyData[^>]*\/>/u, (m) => m + m),
		});
		await expect(decryptOoxmlPackage(dupKeyData, PASSWORD)).rejects.toThrow(/several keyData/u);

		const dupPke = rebuild(await encrypted(), {
			xml: (xml) => xml.replace(/<keyEncryptor uri=[\s\S]*?<\/keyEncryptor>/u, (m) => m + m),
		});
		await expect(decryptOoxmlPackage(dupPke, PASSWORD)).rejects.toThrow(/several password/u);
	});

	it('rejects a missing password key encryptor and malformed XML', async () => {
		const none = rebuild(await encrypted(), {
			xml: (xml) =>
				xml.replace(
					'uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password"',
					'uri="urn:other"',
				),
		});
		await expect(decryptOoxmlPackage(none, PASSWORD)).rejects.toThrow(/no password key encryptor/u);
		const broken = rebuild(await encrypted(), { xml: (xml) => xml.slice(0, 120) });
		await expect(decryptOoxmlPackage(broken, PASSWORD)).rejects.toThrow();
	});
});
