import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import type { Workbook } from '../model.js';
import { loadXlsx } from '../read/index.js';
import { createWorkbook } from '../workbook.js';
import { saveXlsx } from '../write/index.js';
import { verifySheetPassword, verifyWorkbookPassword } from './protection.js';
import { createEditSession } from './session.js';

// Written by Excel 16: `Worksheet.Protect('open sesame')` and `Workbook.Protect('open sesame', True)`.
const SHEET_HASH =
	'algorithmName="SHA-512" hashValue="EQUDCxDLSwlToQ3z0zo+PBWpEhBJ7mifMv+V8jtBNjs+im3SbfefZYbjPpSen/QUg6b03j3NauXpEcjjgn8FDw==" saltValue="sqFVkzKUQmdXc3U0RRSdWA==" spinCount="100000"';
const SHEET_XML = `<sheetProtection ${SHEET_HASH} sheet="1" objects="1" scenarios="1"/>`;
const WORKBOOK_XML =
	'<workbookProtection workbookAlgorithmName="SHA-512" workbookHashValue="4QuxTf3JhGlDJvpkivaz4PJBQQT4eni6s44A5xG6raA+hEc9P0Dm9qt4WzxvBuVxm3FTMmME6hBJmOYbTbTD3g==" workbookSaltValue="bxtG5KBKs/YQVIrtUVrgxw==" workbookSpinCount="100000" lockStructure="1"/>';
const HASH_ATTRIBUTE = /algorithmName|hashValue|saltValue|spinCount|[pP]assword=/;

/** A one-sheet workbook whose parts carry the given protection elements. */
async function protectedBook(sheetXml: string, workbookXml = ''): Promise<Workbook> {
	const zip = await JSZip.loadAsync(await saveXlsx(createWorkbook()));
	const patch = async (name: string, find: string, insert: string) =>
		zip.file(name, (await zip.file(name)!.async('string')).replace(find, `${find}${insert}`));
	await patch('xl/worksheets/sheet1.xml', '<sheetData/>', sheetXml);
	await patch('xl/workbook.xml', '<workbookPr/>', workbookXml);
	return loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
}

/** Saves `workbook` and returns its sheet and workbook XML. */
async function savedXml(workbook: Workbook): Promise<{ sheet: string; book: string }> {
	const zip = await JSZip.loadAsync(await saveXlsx(workbook));
	return {
		sheet: await zip.file('xl/worksheets/sheet1.xml')!.async('string'),
		book: await zip.file('xl/workbook.xml')!.async('string'),
	};
}

const element = (xml: string, name: string) => new RegExp(`<${name}[^>]*/>`).exec(xml)?.[0];

/** The agile hash through node:crypto, independent of `ooxml-core/digest`. */
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

describe('saving sheet protection', () => {
	it('keeps an untouched modern-only protection byte for byte', async () => {
		const { sheet } = await savedXml(await protectedBook(SHEET_XML));
		expect(element(sheet, 'sheetProtection')).toBe(SHEET_XML);
	});

	it('does not invent a spinCount the source left out', async () => {
		const xml = SHEET_XML.replace(' spinCount="100000"', '');
		const book = await protectedBook(xml);
		expect(book.sheets[0]!.protection?.modernHash).not.toHaveProperty('spinCount');
		expect(element((await savedXml(book)).sheet, 'sheetProtection')).toBe(xml);
		const session = createEditSession(book, { recalc: false });
		session.setSheetProtection(0, { sheet: true, allow: ['formatCells'] });
		expect(element((await savedXml(book)).sheet, 'sheetProtection')).not.toContain('spinCount');
	});

	it('writes no hash once the password is removed', async () => {
		const book = await protectedBook(SHEET_XML);
		const session = createEditSession(book, { recalc: false });
		session.setSheetProtection(0, book.sheets[0]!.protection, '');
		const saved = element((await savedXml(book)).sheet, 'sheetProtection');
		expect(saved).toContain('sheet="1"');
		expect(saved).not.toMatch(HASH_ATTRIBUTE);
		const reloaded = await loadXlsx(await saveXlsx(book));
		expect(verifySheetPassword(reloaded.sheets[0], 'anything')).toBe(true);
	});

	it('drops the stale modern hash when the password changes', async () => {
		const book = await protectedBook(SHEET_XML);
		const session = createEditSession(book, { recalc: false });
		session.setSheetProtection(0, book.sheets[0]!.protection, 'new secret');
		const saved = element((await savedXml(book)).sheet, 'sheetProtection') ?? '';
		expect(saved).toMatch(/ password="[0-9A-F]+"/);
		expect(saved).not.toMatch(/algorithmName|hashValue|saltValue|spinCount/);
		const reloaded = (await loadXlsx(await saveXlsx(book))).sheets[0];
		expect(verifySheetPassword(reloaded, 'new secret')).toBe(true);
		expect(verifySheetPassword(reloaded, 'open sesame')).toBe(false);
	});

	it('unprotects only with the right password and then writes no protection', async () => {
		const book = await protectedBook(SHEET_XML);
		const session = createEditSession(book, { recalc: false });
		expect(() => session.setSheetProtection(0, undefined, 'wrong')).toThrow(/not correct/);
		session.setSheetProtection(0, undefined, 'open sesame');
		expect(element((await savedXml(book)).sheet, 'sheetProtection')).toBeUndefined();
	});

	it('unprotects a SHA-256 protected sheet through the edit session', async () => {
		const salt = Buffer.from('0123456789abcdef');
		const hash = nodeHash('secret', salt, 1000, 'sha256');
		const xml = `<sheetProtection algorithmName="SHA-256" hashValue="${hash}" saltValue="${salt.toString('base64')}" spinCount="1000" sheet="1"/>`;
		const book = await protectedBook(xml);
		const session = createEditSession(book, { recalc: false });
		expect(() => session.setSheetProtection(0, undefined, 'Secret')).toThrow(/not correct/);
		expect(book.sheets[0]!.protection?.sheet).toBe(true);
		session.setSheetProtection(0, undefined, 'secret');
		expect(book.sheets[0]!.protection).toBeUndefined();
	});
});

describe('workbook structure protection', () => {
	it('reads the modern hash Excel writes without a legacy password', async () => {
		const book = await protectedBook('', WORKBOOK_XML);
		expect(book.structureLocked).toBe(true);
		expect(book.workbookPasswordHash).toBeUndefined();
		expect(book.workbookModernHash).toMatchObject({ algorithmName: 'SHA-512', spinCount: 100000 });
		expect(verifyWorkbookPassword(book, 'open sesame')).toBe(true);
		expect(verifyWorkbookPassword(book, 'anything')).toBe(false);
		expect(verifyWorkbookPassword(book, '')).toBe(false);
	});

	it('keeps an untouched modern hash on save', async () => {
		const { book } = await savedXml(await protectedBook('', WORKBOOK_XML));
		expect(element(book, 'workbookProtection')).toBe(WORKBOOK_XML);
	});

	it('unlocks only with the right password, and undo restores the hash', async () => {
		const book = await protectedBook('', WORKBOOK_XML);
		const session = createEditSession(book, { recalc: false });
		expect(() => session.setWorkbookProtection(false, 'wrong')).toThrow(/not correct/);
		session.setWorkbookProtection(false, 'open sesame');
		expect(book.workbookModernHash).toBeUndefined();
		expect(element((await savedXml(book)).book, 'workbookProtection')).toBeUndefined();
		session.undo();
		expect(verifyWorkbookPassword(book, 'open sesame')).toBe(true);
		expect(verifyWorkbookPassword(book, 'wrong')).toBe(false);
	});

	it('replaces the modern hash with a new password and removes both with an empty one', async () => {
		const book = await protectedBook('', WORKBOOK_XML);
		const session = createEditSession(book, { recalc: false });
		session.setWorkbookProtection(true, 'new secret');
		let saved = element((await savedXml(book)).book, 'workbookProtection') ?? '';
		expect(saved).toMatch(/workbookPassword="[0-9A-F]+"/);
		expect(saved).not.toMatch(/workbookHashValue|workbookSaltValue|workbookSpinCount/);
		const reloaded = await loadXlsx(await saveXlsx(book));
		expect(verifyWorkbookPassword(reloaded, 'new secret')).toBe(true);
		expect(verifyWorkbookPassword(reloaded, 'open sesame')).toBe(false);
		session.setWorkbookProtection(true, '');
		saved = element((await savedXml(book)).book, 'workbookProtection') ?? '';
		expect(saved).toBe('<workbookProtection lockStructure="1"/>');
	});
});

describe('hashes that cannot be checked', () => {
	const modern = (algorithmName: string, extra: Partial<Record<string, string | number>> = {}) => ({
		algorithmName,
		hashValue:
			'EQUDCxDLSwlToQ3z0zo+PBWpEhBJ7mifMv+V8jtBNjs+im3SbfefZYbjPpSen/QUg6b03j3NauXpEcjjgn8FDw==',
		saltValue: 'sqFVkzKUQmdXc3U0RRSdWA==',
		spinCount: 100000,
		...extra,
	});

	it('never unlocks an unsupported digest without a legacy hash', () => {
		for (const name of ['WHIRLPOOL', 'MD5', 'SHA-3-256']) {
			expect(verifySheetPassword({ sheet: true, modernHash: modern(name) }, ''), name).toBe(false);
			const book = { ...createWorkbook(), structureLocked: true, workbookModernHash: modern(name) };
			expect(verifyWorkbookPassword(book, 'open sesame'), name).toBe(false);
		}
	});

	it('falls back to the legacy hash for an unsupported digest', () => {
		const protection = { sheet: true, modernHash: modern('WHIRLPOOL'), passwordHash: 'CBEB' };
		expect(verifySheetPassword(protection, 'test')).toBe(true);
		expect(verifySheetPassword(protection, 'other')).toBe(false);
	});

	it('rejects malformed base64 and huge spin counts without throwing or stalling', async () => {
		const started = performance.now();
		for (const extra of [{ hashValue: '***' }, { saltValue: 'A' }, { spinCount: 1e12 }]) {
			const protection = { sheet: true, modernHash: modern('SHA-512', extra) };
			expect(verifySheetPassword(protection, 'open sesame')).toBe(false);
		}
		expect(performance.now() - started).toBeLessThan(200);
		const book = await protectedBook(SHEET_XML.replace('hashValue="', 'hashValue="!!'));
		const session = createEditSession(book, { recalc: false });
		expect(() => session.setSheetProtection(0, undefined, 'open sesame')).toThrow(/not correct/);
	});
});
