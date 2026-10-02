import { describe, expect, it } from 'vitest';
import * as common from './common';
import * as react from './react';
import * as vue from './vue';
import * as solid from './solid';
import * as angular from './angular';

describe('surface shared by every published framework package', () => {
	it('re-exports the model API, the element and workbook loading helpers', () => {
		for (const name of [
			'createWorkbook',
			'loadXlsx',
			'saveXlsx',
			'loadWorkbook',
			'detectWorkbookFormat',
			'loadLegacyXls',
			'csvToWorkbook',
			'sheetToCsv',
			'saveWorkbook',
			'defineXlsxEditor',
			'XlsxEditorElement',
			'normalizeEditorLocale',
		] as const)
			expect(common[name], name).toBeTypeOf('function');
		expect(common.XLSX_EDITOR_EVENTS.length).toBeGreaterThan(0);
	});

	it('exports each framework component under the documented name', () => {
		expect(react.SpreadsheetEditor).toBeDefined();
		expect(vue.SpreadsheetEditor).toBeDefined();
		expect(solid.SpreadsheetEditor).toBeTypeOf('function');
		expect(angular.SpreadsheetEditorComponent).toBeTypeOf('function');
	});
});
