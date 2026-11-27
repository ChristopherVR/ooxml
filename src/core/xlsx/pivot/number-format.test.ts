import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { createEditSession } from '../edit/session';

it('writes and reads back the number formats of PivotTable value fields', async () => {
	const wb = await loadXlsx(
		readFileSync(new URL('../__fixtures__/excel-pivot-values.xlsx', import.meta.url)),
	);
	const edit = createEditSession(wb);
	const summary = wb.sheets.findIndex((sheet) => sheet.name === 'Summary');
	const pivot = wb.sheets[summary]!.pivotTables![0]!;
	edit.refreshPivotTable(summary, pivot.name, {
		valueNumberFormat: '#,##0',
		additionalValues: [
			{ field: 'Quantity', aggregate: 'count', caption: 'Count of Quantity' },
			{
				field: 'Amount',
				aggregate: 'average',
				caption: 'Average of Amount',
				numberFormat: '"$"#,##0.0',
			},
		],
	});
	const bytes = await saveXlsx(wb);
	const zip = await JSZip.loadAsync(bytes);
	const part = Object.keys(zip.files).find((path) =>
		/^xl\/pivotTables\/pivotTable\d+\.xml$/.test(path),
	)!;
	const table = await zip.file(part)!.async('string');
	const fields = [...table.matchAll(/<dataField\b[^>]*>/g)].map((match) => match[0]);
	expect(fields[0]).toContain('numFmtId="3"');
	expect(fields[1]).not.toContain('numFmtId');
	const custom = /numFmtId="(\d+)"/.exec(fields[2] ?? '')?.[1];
	expect(Number(custom)).toBeGreaterThanOrEqual(164);
	const styles = await zip.file('xl/styles.xml')!.async('string');
	expect(styles).toContain(`<numFmt numFmtId="${custom}" formatCode="&quot;$&quot;#,##0.0"/>`);
	const reopened = (await loadXlsx(bytes)).sheets[summary]!.pivotTables![0]!;
	expect(reopened.valueNumberFormat).toBe('#,##0');
	expect(reopened.additionalValues?.[0]?.numberFormat).toBeUndefined();
	expect(reopened.additionalValues?.[1]?.numberFormat).toBe('"$"#,##0.0');
	edit.refreshPivotTable(summary, pivot.name, { valueNumberFormat: null });
	expect(wb.sheets[summary]!.pivotTables![0]!.valueNumberFormat).toBeUndefined();
});
