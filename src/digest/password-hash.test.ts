import { createHash, randomBytes } from 'node:crypto';
import { decodeBase64, encodeBase64 } from './base64.js';
import { MAX_SPIN_COUNT, hashPassword, verifyPasswordHash, type PasswordHash } from './index.js';

// Written by Excel 16 for `Worksheet.Protect('open sesame')`.
const EXCEL_SHEET: PasswordHash = {
	algorithmName: 'SHA-512',
	hashValue:
		'EQUDCxDLSwlToQ3z0zo+PBWpEhBJ7mifMv+V8jtBNjs+im3SbfefZYbjPpSen/QUg6b03j3NauXpEcjjgn8FDw==',
	saltValue: 'sqFVkzKUQmdXc3U0RRSdWA==',
	spinCount: 100000,
};

// Written by Excel 16 for `Workbook.Protect('open sesame', True)` (workbookHashValue and so on).
const EXCEL_WORKBOOK: PasswordHash = {
	algorithmName: 'SHA-512',
	hashValue:
		'4QuxTf3JhGlDJvpkivaz4PJBQQT4eni6s44A5xG6raA+hEc9P0Dm9qt4WzxvBuVxm3FTMmME6hBJmOYbTbTD3g==',
	saltValue: 'bxtG5KBKs/YQVIrtUVrgxw==',
	spinCount: 100000,
};

/** An independent agile hash on node:crypto, for the digests Excel does not write. */
function nodeHash(password: string, salt: Buffer, spinCount: number, algorithm: string): string {
	let value = createHash(algorithm)
		.update(Buffer.concat([salt, Buffer.from(password, 'utf16le')]))
		.digest();
	for (let i = 0; i < spinCount; i++) {
		const round = Buffer.alloc(4);
		round.writeUInt32LE(i);
		value = createHash(algorithm)
			.update(Buffer.concat([value, round]))
			.digest();
	}
	return value.toString('base64');
}

describe('hashPassword', () => {
	it('reproduces the sheet and workbook hashes Excel 16 wrote', () => {
		expect(hashPassword('open sesame', EXCEL_SHEET)).toBe(EXCEL_SHEET.hashValue);
		expect(hashPassword('open sesame', EXCEL_WORKBOOK)).toBe(EXCEL_WORKBOOK.hashValue);
	});

	it.each([
		['SHA-1', 'sha1'],
		['SHA256', 'sha256'],
		['sha_384', 'sha384'],
		['SHA-512', 'sha512'],
	])('matches node:crypto for %s', (algorithmName, node) => {
		const salt = randomBytes(16);
		const options = { algorithmName, saltValue: salt.toString('base64'), spinCount: 1000 };
		expect(hashPassword('päss 中', options)).toBe(nodeHash('päss 中', salt, 1000, node));
	});

	it('treats a missing spin count as no extra rounds', () => {
		const salt = randomBytes(16);
		expect(
			hashPassword('x', { algorithmName: 'SHA-256', saltValue: salt.toString('base64') }),
		).toBe(nodeHash('x', salt, 0, 'sha256'));
	});

	it('returns undefined for unsupported digests and unusable parameters', () => {
		const base = { saltValue: 'eA==', spinCount: 1 };
		expect(hashPassword('x', { ...base, algorithmName: 'WHIRLPOOL' })).toBeUndefined();
		expect(hashPassword('x', { ...base, algorithmName: 'MD5' })).toBeUndefined();
		expect(hashPassword('x', { ...base, algorithmName: 'SHA-1', saltValue: '!!' })).toBeUndefined();
		for (const spinCount of [-1, 1.5, Number.NaN, Infinity, MAX_SPIN_COUNT + 1, 2 ** 32])
			expect(
				hashPassword('x', { ...base, algorithmName: 'SHA-1', spinCount }),
				String(spinCount),
			).toBeUndefined();
	});
});

describe('verifyPasswordHash', () => {
	it('accepts only the password Excel used', () => {
		expect(verifyPasswordHash('open sesame', EXCEL_SHEET)).toBe(true);
		expect(verifyPasswordHash('open sesam', EXCEL_SHEET)).toBe(false);
		expect(verifyPasswordHash('', EXCEL_SHEET)).toBe(false);
		expect(verifyPasswordHash('open sesame', EXCEL_WORKBOOK)).toBe(true);
		expect(verifyPasswordHash('Open sesame', EXCEL_WORKBOOK)).toBe(false);
	});

	it('compares decoded bytes, tolerating white space and missing padding', () => {
		const unpadded = EXCEL_SHEET.hashValue.replace(/=+$/, '');
		const spaced = `${unpadded.slice(0, 40)}\n  ${unpadded.slice(40)}`;
		expect(verifyPasswordHash('open sesame', { ...EXCEL_SHEET, hashValue: spaced })).toBe(true);
		expect(
			verifyPasswordHash('open sesame', { ...EXCEL_SHEET, saltValue: 'sqFVkzKUQmdXc3U0RRSdWA' }),
		).toBe(true);
	});

	it('returns false, never throws, for malformed base64', () => {
		for (const hashValue of ['', '***', 'A', 'AB=C', '=AAA'])
			expect(verifyPasswordHash('open sesame', { ...EXCEL_SHEET, hashValue }), hashValue).toBe(
				false,
			);
		expect(verifyPasswordHash('open sesame', { ...EXCEL_SHEET, saltValue: '%%' })).toBe(false);
	});

	it('returns false at once for a huge or invalid spin count', () => {
		const started = performance.now();
		for (const spinCount of [MAX_SPIN_COUNT + 1, 1e12, Infinity, Number.NaN, -5])
			expect(verifyPasswordHash('open sesame', { ...EXCEL_SHEET, spinCount })).toBe(false);
		expect(performance.now() - started).toBeLessThan(100);
	});

	it('returns undefined for digests it does not compute', () => {
		expect(verifyPasswordHash('x', { ...EXCEL_SHEET, algorithmName: 'WHIRLPOOL' })).toBeUndefined();
		expect(verifyPasswordHash('x', { ...EXCEL_SHEET, algorithmName: 'MD5' })).toBeUndefined();
	});
});

describe('base64', () => {
	it('round-trips every length against Buffer', () => {
		for (let length = 0; length < 40; length++) {
			const bytes = new Uint8Array(randomBytes(length));
			const encoded = encodeBase64(bytes);
			expect(encoded).toBe(Buffer.from(bytes).toString('base64'));
			expect(decodeBase64(encoded)).toEqual(bytes);
			expect(decodeBase64(encoded.replace(/=+$/, ''))).toEqual(bytes);
		}
	});

	it('rejects malformed input', () => {
		for (const text of ['A', 'AAAAA', 'AB=C', 'AA=', 'A===', 'AA*A', '=AAA'])
			expect(decodeBase64(text), text).toBeUndefined();
	});
});
