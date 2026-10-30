import JSZip from 'jszip';
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { NS } from '../../xml/index';
import { parseRange } from '../address';
import { createEditSession } from '../edit/session';
import { listValidationOptions, validateCellInput } from '../edit/validation';
import { loadXlsx } from '../read/index';
import { saveXlsx } from './index';
import { miniPackage, roundTrip, ws } from './mini-package-fixtures';

const URI = '{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}';
const XM = 'http://schemas.microsoft.com/office/excel/2006/main';
// Structure recorded from Excel 16: a list validation with a direct reference to another sheet.
const rule = `<x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorTitle="Pick a colour" error="Choose from the list." imeMode="hiragana" xmlns:xr="${NS.xr}" xr:uid="{11111111-1111-1111-1111-111111111111}"><x14:formula1><xm:f>Lookup!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>B2:B5</xm:sqref></x14:dataValidation>`;
const extensions = `<extLst><ext uri="${URI}" xmlns:x14="${NS.x14}"><x14:dataValidations count="1" disablePrompts="1" xWindow="12" xmlns:xm="${XM}">${rule}</x14:dataValidations></ext><ext uri="other"><custom xmlns="urn:keep" value="unchanged"/></ext></extLst>`;
const original = () =>
	miniPackage({
		sheets: [
			{ name: 'Input', xml: ws('', '', extensions) },
			{
				name: 'Lookup',
				xml: ws(
					['red', 'green', 'blue']
						.map(
							(value, i) =>
								`<row r="${i + 1}"><c r="A${i + 1}" t="inlineStr"><is><t>${value}</t></is></c></row>`,
						)
						.join(''),
				),
			},
		],
	});
const xmlOf = async (bytes: Uint8Array) =>
	(await JSZip.loadAsync(bytes)).file('xl/worksheets/sheet1.xml')!.async('string');

describe('Excel 2010 data validations', () => {
	it('edits a native Excel rule and rewrites it after renaming and inserting source rows', async () => {
		const workbook = await loadXlsx(
			new Uint8Array(
				readFileSync(new URL('../__fixtures__/excel-x14-validation.xlsx', import.meta.url)),
			),
		);
		const input = workbook.sheets.findIndex((sheet) => sheet.name === 'Input');
		const lookup = workbook.sheets.findIndex((sheet) => sheet.name === 'Lookup');
		expect(listValidationOptions(workbook, input, 1, 1)).toEqual(['red', 'green', 'blue']);
		const edit = createEditSession(workbook);
		edit.renameSheet(lookup, 'New Colours');
		edit.insertRows(lookup, 0, 1);
		const reloaded = await loadXlsx(await saveXlsx(workbook));
		expect(reloaded.sheets[input]!.dataValidations[0]!.formula1).toBe("'New Colours'!$A$2:$A$4");
		expect(listValidationOptions(reloaded, input, 1, 1)).toEqual(['red', 'green', 'blue']);
	});

	it('retains surviving rules when clearing part of a validation without duplicating revision ids', async () => {
		const workbook = await loadXlsx(await original());
		createEditSession(workbook).setDataValidation(0, undefined, parseRange('B3')!);
		const bytes = await saveXlsx(workbook);
		const reloaded = await loadXlsx(bytes);
		expect(listValidationOptions(reloaded, 0, 1, 1)).toEqual(['red', 'green', 'blue']);
		expect(listValidationOptions(reloaded, 0, 2, 1)).toBeUndefined();
		expect(listValidationOptions(reloaded, 0, 3, 1)).toEqual(['red', 'green', 'blue']);
		expect((await xmlOf(bytes)).match(/xr:uid=/g)).toHaveLength(1);
	});

	it('loads cross-sheet list options, messages and restrictions through the existing validation engine', async () => {
		const workbook = await loadXlsx(await original());
		expect(workbook.sheets[0]!.dataValidations).toHaveLength(1);
		expect(workbook.sheets[0]!.dataValidations[0]).toMatchObject({
			type: 'list',
			formula1: 'Lookup!$A$1:$A$3',
			ranges: [parseRange('B2:B5')],
			showDropDown: true,
		});
		expect(listValidationOptions(workbook, 0, 1, 1)).toEqual(['red', 'green', 'blue']);
		expect(validateCellInput(workbook, 0, 1, 1, 'red').ok).toBe(true);
		expect(validateCellInput(workbook, 0, 1, 1, 'orange').ok).toBe(false);
	});

	it('keeps copied rules and sheets without duplicating revision identities or mutating source metadata', async () => {
		const workbook = await loadXlsx(await original());
		const edit = createEditSession(workbook);
		const current = workbook.sheets[0]!.dataValidations[0]!;
		edit.setDataValidation(0, { ...current, formula1: 'Lookup!$A$2:$A$3' }, parseRange('B2')!);
		edit.duplicateSheet(0);
		const before = structuredClone(workbook.sheets.map((sheet) => sheet.dataValidations));
		const bytes = await saveXlsx(workbook);
		const zip = await JSZip.loadAsync(bytes);
		const xml = (
			await Promise.all(
				zip.file(/^xl\/worksheets\/sheet\d+\.xml$/).map((part) => part.async('string')),
			)
		).join('');
		expect(xml.match(/xr:uid=/g)).toHaveLength(1);
		expect(workbook.sheets.map((sheet) => sheet.dataValidations)).toEqual(before);
		const reloaded = await loadXlsx(bytes);
		for (const index of [0, 1]) {
			expect(listValidationOptions(reloaded, index, 1, 1)).toEqual(['green', 'blue']);
			expect(listValidationOptions(reloaded, index, 2, 1)).toEqual(['red', 'green', 'blue']);
		}
	});

	it('writes edited formulas once, keeps unrelated XML and collection settings, and reloads', async () => {
		const { again, part } = await roundTrip(await original(), (workbook) => {
			const current = workbook.sheets[0]!.dataValidations[0]!;
			createEditSession(workbook).setDataValidation(
				0,
				{ ...current, formula1: 'Lookup!$A$2:$A$3', error: 'Updated error' },
				parseRange('B2:B5')!,
			);
		});
		expect(listValidationOptions(again, 0, 1, 1)).toEqual(['green', 'blue']);
		const xml = (await part('xl/worksheets/sheet1.xml'))!;
		expect(xml.match(/<x14:dataValidation\s/g)).toHaveLength(1);
		expect(xml).not.toMatch(/<dataValidations[\s>]/);
		for (const text of [
			'imeMode="hiragana"',
			'disablePrompts="1"',
			'xWindow="12"',
			'Updated error',
			'value="unchanged"',
		])
			expect(xml).toContain(text);
	});

	it('edits both operands and flags through the shared codec when changing the validation type', async () => {
		const { again } = await roundTrip(await original(), (workbook) => {
			const current = workbook.sheets[0]!.dataValidations[0]!;
			createEditSession(workbook).setDataValidation(
				0,
				{
					...current,
					type: 'whole',
					operator: 'between',
					formula1: '1',
					formula2: '10',
					showDropDown: false,
					showInputMessage: false,
					errorStyle: 'warning',
				},
				parseRange('B2:B5')!,
			);
		});
		expect(again.sheets[0]!.dataValidations[0]).toMatchObject({
			type: 'whole',
			operator: 'between',
			formula1: '1',
			formula2: '10',
			showDropDown: false,
			errorStyle: 'warning',
		});
		expect(again.sheets[0]!.dataValidations[0]!.showInputMessage).toBeFalsy();
		expect(validateCellInput(again, 0, 1, 1, 5).ok).toBe(true);
		expect(validateCellInput(again, 0, 1, 1, 11).ok).toBe(false);
	});

	it('removes cleared rules without resurrecting the carried extension, and supports undo', async () => {
		const workbook = await loadXlsx(await original());
		const edit = createEditSession(workbook);
		edit.setDataValidation(0, undefined, parseRange('B2:B5')!);
		const bytes = await saveXlsx(workbook);
		expect(await xmlOf(bytes)).not.toContain(URI);
		expect(await xmlOf(bytes)).toContain('value="unchanged"');
		expect((await loadXlsx(bytes)).sheets[0]!.dataValidations).toEqual([]);
		edit.undo();
		expect((await loadXlsx(await saveXlsx(workbook))).sheets[0]!.dataValidations).toHaveLength(1);
	});

	it('writes newly authored cross-sheet references as x14, while literal exclamation marks stay in the base codec', async () => {
		const workbook = await loadXlsx(
			await miniPackage({ sheets: [{ name: 'Input', xml: ws('') }] }),
		);
		const edit = createEditSession(workbook);
		edit.setDataValidation(
			0,
			{ ranges: [], type: 'list', formula1: "'Lookup Sheet'!$A$1:$A$3" },
			parseRange('B2')!,
		);
		edit.setDataValidation(
			0,
			{ ranges: [], type: 'list', formula1: '"yes!,no"' },
			parseRange('C2')!,
		);
		const bytes = await saveXlsx(workbook);
		const xml = await xmlOf(bytes);
		expect(xml).toContain('<dataValidations count="1">');
		expect(xml).toContain('<x14:dataValidations');
		expect(xml).toContain('<xm:sqref>B2</xm:sqref>');
		expect((await loadXlsx(bytes)).sheets[0]!.dataValidations).toHaveLength(2);
	});
});
