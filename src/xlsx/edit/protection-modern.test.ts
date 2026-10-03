import { NS, parseXml, type XmlElement } from '../../xml/index.js';
import { readProtection } from '../read/sheet-props.js';
import { modernPasswordHash, verifySheetPassword, verifySheetPasswordAsync } from './protection.js';

// Written by Excel 16 for `Worksheet.Protect('open sesame')`: the modern hash only, no legacy one.
const EXCEL_XML =
	'<sheetProtection algorithmName="SHA-512" hashValue="EQUDCxDLSwlToQ3z0zo+PBWpEhBJ7mifMv+V8jtBNjs+im3SbfefZYbjPpSen/QUg6b03j3NauXpEcjjgn8FDw==" saltValue="sqFVkzKUQmdXc3U0RRSdWA==" spinCount="100000" sheet="1" objects="1" scenarios="1"/>';

function excelProtection() {
	const root = parseXml(`<x xmlns="${NS.x}">${EXCEL_XML}</x>`).documentElement;
	const protection = readProtection(root.firstChild as XmlElement);
	if (!protection) throw new Error('no protection read');
	return protection;
}

describe('modern (SHA-512) sheet protection', () => {
	it('reads the modern hash into the model', () => {
		expect(excelProtection().modernHash).toEqual({
			algorithmName: 'SHA-512',
			hashValue:
				'EQUDCxDLSwlToQ3z0zo+PBWpEhBJ7mifMv+V8jtBNjs+im3SbfefZYbjPpSen/QUg6b03j3NauXpEcjjgn8FDw==',
			saltValue: 'sqFVkzKUQmdXc3U0RRSdWA==',
			spinCount: 100000,
		});
	});

	it('checks the SHA-512 hash synchronously', () => {
		const protection = excelProtection();
		expect(verifySheetPassword(protection, 'open sesame')).toBe(true);
		expect(verifySheetPassword(protection, 'anything')).toBe(false);
		expect(verifySheetPassword(protection, '')).toBe(false);
	});

	it('verifies the password Excel used, and only that one', async () => {
		const protection = excelProtection();
		await expect(verifySheetPasswordAsync(protection, 'open sesame')).resolves.toBe(true);
		await expect(verifySheetPasswordAsync(protection, 'open sesam')).resolves.toBe(false);
		await expect(verifySheetPasswordAsync(protection, '')).resolves.toBe(false);
	});

	it('reproduces the hash Excel wrote', async () => {
		const hash = excelProtection().modernHash;
		if (!hash) throw new Error('no modern hash');
		await expect(modernPasswordHash('open sesame', hash)).resolves.toBe(hash.hashValue);
	});

	it('unlocks a protected sheet that has no password at all', async () => {
		const protection = { sheet: true };
		expect(verifySheetPassword(protection, '')).toBe(true);
		await expect(verifySheetPasswordAsync(protection, '')).resolves.toBe(true);
	});

	it('does not accept any password for an unknown digest without a legacy hash', async () => {
		const protection = {
			sheet: true,
			modernHash: { algorithmName: 'WHIRLPOOL', hashValue: 'x', saltValue: 'eA==', spinCount: 1 },
		};
		await expect(verifySheetPasswordAsync(protection, 'anything')).resolves.toBe(false);
	});
});
