import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { saveXlsx } from '../write/index';
import { loadXlsx } from './index';

const fixture = () =>
	new Uint8Array(
		readFileSync(path.join(import.meta.dirname, '..', '__fixtures__', 'excel-custom-views.xlsx')),
	);

const COMPACT = '{9EDFE334-6832-4746-8699-F0DD94A8B14B}';
const FULL = '{7CAA533B-730A-4C0B-9CBD-DDDF4C014020}';

describe('custom views in a file', () => {
	it('reads the views Excel wrote', async () => {
		const wb = await loadXlsx(fixture());
		expect(wb.customViews?.map((v) => [v.name, v.guid])).toEqual([
			['Full', FULL],
			['Compact', COMPACT],
		]);
		expect(wb.customViews?.[0]).toMatchObject({
			includePrintSettings: false,
			includeHiddenRowCol: false,
		});
		expect(wb.customViews?.[1]).toMatchObject({
			includePrintSettings: true,
			includeHiddenRowCol: true,
		});
		const data = wb.sheets[0]!;
		const compact = data.customViews!.find((v) => v.guid === COMPACT)!;
		expect(compact.view).toMatchObject({ zoom: 85, showGridLines: false, showHeaders: true });
		expect(compact.hiddenColumns).toBe(true);
		expect(compact.showAutoFilter).toBe(true);
		expect(compact.filter?.range).toEqual({
			start: { row: 0, col: 0 },
			end: { row: 5, col: 3 },
		});
		expect(compact.filter?.columns?.[0]).toMatchObject({ offset: 1, values: ['R2C2'] });
		expect(compact.pageSetup).toMatchObject({ orientation: 'landscape', header: '&CCompact' });
		const full = data.customViews!.find((v) => v.guid === FULL)!;
		expect(full.view.zoom).toBe(100);
		expect(full.filter).toBeUndefined();
		expect(wb.sheets[1]!.customViews).toHaveLength(2);
		// The hidden-column set lives in a hidden defined name.
		expect(
			wb.definedNames.find((n) => n.name === 'Z_9EDFE334_6832_4746_8699_F0DD94A8B14B_.wvu.Cols')
				?.formula,
		).toBe('Data!$C:$C');
	});

	it('writes them back and reads the same views again', async () => {
		const wb = await loadXlsx(fixture());
		const zip = await JSZip.loadAsync(await saveXlsx(wb));
		const book = await zip.file('xl/workbook.xml')!.async('string');
		expect(book).toContain('<customWorkbookViews>');
		expect(book).toContain('includePrintSettings="0"');
		expect(book).toContain('wvu.Cols');
		const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
		expect(sheet).toContain(`<customSheetView guid="${COMPACT}"`);
		expect(sheet).toContain('scale="85"');
		expect(sheet).toContain('showGridLines="0"');
		expect(sheet).toContain('hiddenColumns="1"');
		const again = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
		expect(again.customViews).toEqual(wb.customViews);
		expect(again.sheets[0]!.customViews).toEqual(wb.sheets[0]!.customViews);
		expect(again.sheets[1]!.customViews).toEqual(wb.sheets[1]!.customViews);
	});
});
