import { describe, expect, it } from 'vitest';
import * as core from './index';
import * as load from './load';

// Entry contract: the published names stay stable even though the code lives in ooxml-core.
describe('@christophervr/xlsx-core entry', () => {
	it('re-exports the workbook model and package codec', () => {
		for (const name of [
			'createWorkbook',
			'createWorksheet',
			'parseAddress',
			'formatAddress',
			'getCell',
			'putCell',
			'loadXlsx',
			'saveXlsx',
		] as const)
			expect(core[name], name).toBeTypeOf('function');
	});

	it('keeps the legacy .xls reader out of the main entry', () => {
		expect('loadLegacyXls' in core).toBe(false);
	});

	it('re-exports format detection and loading from /load', () => {
		for (const name of [
			'loadWorkbook',
			'detectWorkbookFormat',
			'loadLegacyXls',
			'parseCsv',
			'csvToWorkbook',
			'sheetToCsv',
			'saveWorkbook',
		] as const)
			expect(load[name], name).toBeTypeOf('function');
	});

	it('round-trips a new workbook through the main entry', async () => {
		const workbook = core.createWorkbook();
		const bytes = await core.saveXlsx(workbook);
		expect(load.detectWorkbookFormat(bytes)).toBe('xlsx');
		const reopened = await core.loadXlsx(bytes);
		expect(reopened.sheets.length).toBe(workbook.sheets.length);
	});
});
