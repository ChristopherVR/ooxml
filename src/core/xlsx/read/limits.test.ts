import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { getCell, putCell } from '../cells';
import { loadWorkbook, saveWorkbook } from '../load/detect';
import { createWorkbook } from '../workbook';
import { saveXlsx } from '../write/index';
import * as xlsxEntry from '../index';
import * as loadEntry from '../load/index';
import { ZipLimitError } from '../../opc/index';
import { LiteDeclinedError } from '../../xml/index';
import { isWorkbookTooLargeError, WorkbookTooLargeError } from './limits';
import { loadXlsx } from './load';
import { splitSheetData } from './sheet-data-source';

const STRINGS = 200;

/** A small synthetic workbook: distinct strings (a shared-strings part) and numbers. */
async function workbookBytes(): Promise<Uint8Array> {
	const wb = createWorkbook();
	const sheet = wb.sheets[0];
	if (!sheet) throw new Error('no sheet');
	for (let row = 0; row < STRINGS; row++) {
		putCell(sheet, row, 0, { value: `label ${row}` });
		putCell(sheet, row, 1, { value: row * 1.5 });
	}
	return saveXlsx(wb);
}

/** Entry count, largest part and total inflated size of a package. */
async function measure(bytes: Uint8Array) {
	const zip = await JSZip.loadAsync(bytes);
	const sizes = await Promise.all(
		Object.values(zip.files)
			.filter((entry) => !entry.dir)
			.map(async (entry) => (await entry.async('uint8array')).byteLength),
	);
	return {
		parts: sizes.length,
		largest: Math.max(...sizes),
		total: sizes.reduce((sum, size) => sum + size, 0),
	};
}

const tooLarge = (limit: string, maximum: number) =>
	expect.objectContaining({
		name: 'WorkbookTooLargeError',
		code: 'workbook-too-large',
		limit,
		maximum,
	});

describe('loadXlsx limits', () => {
	it('names the limit a package exceeds and loads it once that limit allows it', async () => {
		const bytes = await workbookBytes();
		const { parts, largest, total } = await measure(bytes);
		const exact = {
			maxInputBytes: bytes.byteLength,
			maxParts: parts,
			maxPartBytes: largest,
			maxTotalBytes: total,
			// The shared strings: the root and an `<si><t>` pair per string.
			maxXmlElements: 1 + 2 * STRINGS,
		};
		for (const [limit, maximum] of Object.entries(exact)) {
			await expect(loadXlsx(bytes, { limits: { [limit]: maximum - 1 } })).rejects.toEqual(
				tooLarge(limit, maximum - 1),
			);
			const wb = await loadXlsx(bytes, { limits: { [limit]: maximum } });
			expect(getCell(wb.sheets[0]!, STRINGS - 1, 0)?.value).toBe(`label ${STRINGS - 1}`);
		}
	});

	/** The workbook with its first worksheet part edited, and that part's new XML. */
	async function withSheet(edit: (xml: string) => string) {
		const zip = await JSZip.loadAsync(await workbookBytes());
		const part = 'xl/worksheets/sheet1.xml';
		const xml = edit(await zip.file(part)!.async('string'));
		zip.file(part, xml);
		return { xml, bytes: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }) };
	}

	it('applies the element limit to a worksheet read as a whole', async () => {
		// The fast cell reader declines a processing instruction in the cell data, so the whole part
		// goes to the XML parser.
		const { xml, bytes } = await withSheet((part) =>
			part.replace('<sheetData>', '<sheetData><?keep?>'),
		);
		expect(() => [...splitSheetData(xml)!.rows]).toThrow(LiteDeclinedError);
		const limits = { maxXmlElements: 1000 };
		await expect(loadXlsx(bytes, { limits })).rejects.toEqual(
			expect.objectContaining({ message: expect.stringMatching(/^XLSX worksheet XML/) }),
		);
		await expect(loadXlsx(bytes, { limits })).rejects.toEqual(tooLarge('maxXmlElements', 1000));
		const wb = await loadXlsx(bytes, { limits: { maxXmlElements: 10_000 } });
		expect(getCell(wb.sheets[0]!, 3, 1)?.value).toBe(4.5);
	});

	it('does not count cell data the fast reader takes, also in a part with comments', async () => {
		const { xml, bytes } = await withSheet((part) =>
			part.replace('<sheetData>', '<!--x--><sheetData>'),
		);
		expect(splitSheetData(xml)).toBeDefined();
		const wb = await loadXlsx(bytes, { limits: { maxXmlElements: 1000 } });
		expect(getCell(wb.sheets[0]!, 3, 1)?.value).toBe(4.5);
	});

	it('keeps the default limits', async () => {
		const oversized = new Uint8Array(51 * 1024 * 1024);
		await expect(loadXlsx(oversized)).rejects.toEqual(tooLarge('maxInputBytes', 50 * 1024 * 1024));
		// Raised, the size check passes and the bytes are read (and refused as not a package).
		await expect(
			loadXlsx(oversized, { limits: { maxInputBytes: 64 * 1024 * 1024 } }),
		).rejects.toThrow(/Not a valid XLSX package/);
	});

	it('rejects a limit that is not a positive finite number', async () => {
		const bytes = await workbookBytes();
		for (const value of [0, -1, Number.NaN, Number.POSITIVE_INFINITY])
			await expect(loadXlsx(bytes, { limits: { maxParts: value } })).rejects.toThrow(RangeError);
	});

	it('keeps reporting unsafe part names as before', async () => {
		const zip = await JSZip.loadAsync(await workbookBytes());
		zip.file('../outside.xml', '<a/>');
		const bytes = await zip.generateAsync({ type: 'uint8array' });
		await expect(loadXlsx(bytes, { limits: { maxParts: 1000 } })).rejects.toMatchObject({
			name: 'ZipLimitError',
		});
	});
});

describe('loadWorkbook limits', () => {
	it('passes the limits on, also for a password-protected package', async () => {
		const bytes = await workbookBytes();
		const { parts } = await measure(bytes);
		const limits = { maxParts: parts - 1 };
		await expect(loadWorkbook(bytes, { limits })).rejects.toEqual(tooLarge('maxParts', parts - 1));
		const workbook = await loadXlsx(bytes);
		const encrypted = await saveWorkbook(workbook, 'xlsx', {
			password: 'secret',
			encryption: { spinCount: 1000 },
		});
		await expect(loadWorkbook(encrypted, { password: 'secret', limits })).rejects.toEqual(
			tooLarge('maxParts', parts - 1),
		);
		const opened = await loadWorkbook(encrypted, {
			password: 'secret',
			limits: { maxParts: parts },
		});
		expect(getCell(opened.sheets[0]!, 0, 0)?.value).toBe('label 0');
	});

	it('rejects a bad limit before detecting or decrypting the input', async () => {
		const bad = { limits: { maxParts: Number.NaN } };
		const csv = new TextEncoder().encode('a,b\n1,2\n');
		await expect(loadWorkbook(csv, { ...bad, fileName: 'data.csv' })).rejects.toThrow(RangeError);
		const encrypted = await saveWorkbook(await loadXlsx(await workbookBytes()), 'xlsx', {
			password: 'secret',
			encryption: { spinCount: 1000 },
		});
		// A wrong password would fail later, in decryption: the limit is checked first.
		await expect(loadWorkbook(encrypted, { ...bad, password: 'wrong' })).rejects.toThrow(
			RangeError,
		);
	});
});

describe('isWorkbookTooLargeError', () => {
	it('recognises the error by its code, also a copy from another entry point', async () => {
		const error = await loadXlsx(new Uint8Array(51 * 1024 * 1024)).catch((e: unknown) => e);
		expect(isWorkbookTooLargeError(error)).toBe(true);
		// Each entry point bundles its own class: same code, another constructor.
		const copy = Object.assign(new Error('XLSX package exceeds the 1 MiB limit'), {
			code: 'workbook-too-large',
			limit: 'maxInputBytes',
			maximum: 1,
		});
		expect(copy instanceof WorkbookTooLargeError).toBe(false);
		expect(isWorkbookTooLargeError(copy)).toBe(true);
		const others = [
			new RangeError('limits.maxParts must be a positive finite number'),
			new ZipLimitError('unsafe part name', { limit: 1, kind: 'partName' }),
			{ code: 'workbook-too-large' },
			'workbook-too-large',
			undefined,
		];
		for (const other of others) expect(isWorkbookTooLargeError(other)).toBe(false);
	});

	it('is exported with the error by both entry points', () => {
		for (const entry of [xlsxEntry, loadEntry]) {
			expect(entry.isWorkbookTooLargeError).toBe(isWorkbookTooLargeError);
			expect(entry.WorkbookTooLargeError).toBe(WorkbookTooLargeError);
		}
	});
});
