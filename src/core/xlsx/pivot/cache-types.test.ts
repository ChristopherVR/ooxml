import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { NS, children, first, parseXml } from '../../xml/index';
import { createWorkbook } from '../workbook';
import { createEditSession } from '../edit/session';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { pivotGrid } from './aggregate';

for (const type of ['blank', 'boolean', 'error'] as const)
	it(`writes native-compatible shared-item flags and labels for ${type} groups`, async () => {
		const wb = createWorkbook({ sheets: ['Data', 'Summary'] });
		const edit = createEditSession(wb);
		const outer = type === 'blank' ? '' : type === 'boolean' ? 'TRUE' : '=NA()';
		[
			['Category', 'Region', 'Amount'],
			[outer, 'East', '10'],
			['Veg', 'West', '20'],
		].forEach((line, r) => line.forEach((value, c) => edit.setCellInput(0, r, c, value)));
		edit.createPivotTable(
			1,
			{ row: 0, col: 0 },
			{ source: 'Data!A1:C3', rowField: 'Category', columnField: 'Region', valueField: 'Amount' },
		);
		const values = pivotGrid(wb.sheets[1]!.pivotTables![0]!).values;
		const label = type === 'blank' ? '(blank)' : type === 'boolean' ? 'TRUE' : '#N/A';
		expect(values.some((line) => line[0] === label)).toBe(true);
		const bytes = await saveXlsx(wb);
		const zip = await JSZip.loadAsync(bytes);
		const xml = await zip.file('xl/pivotCache/pivotCacheDefinition1.xml')!.async('string');
		const field = children(
			first(parseXml(xml).documentElement, 'cacheFields', NS.x)!,
			'cacheField',
			NS.x,
		)[0]!;
		const items = first(field, 'sharedItems', NS.x)!;
		expect(items.getAttribute('containsString')).toBe('1');
		expect(items.getAttribute('containsMixedTypes')).toBe('0');
		expect(items.getAttribute('containsBlank')).toBe(type === 'blank' ? '1' : '0');
		const reloaded = await loadXlsx(bytes);
		expect(reloaded.warnings).toEqual([]);
		expect(pivotGrid(reloaded.sheets[1]!.pivotTables![0]!).values).toEqual(values);
	});
