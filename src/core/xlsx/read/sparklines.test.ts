import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { parseRange } from '../address';
import { createEditSession } from '../edit/index';
import { createWorkbook } from '../workbook';
import { saveXlsx } from '../write/index';
import { loadXlsx } from './index';
import { readSparklineGroups, sheetSparklineGroups } from './sparklines';

const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const XM = 'http://schemas.microsoft.com/office/excel/2006/main';
const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

// Authored by Excel 16 through COM (generate-excel-sparklines.ps1).
const fixture = () =>
	new Uint8Array(
		readFileSync(path.join(import.meta.dirname, '..', '__fixtures__', 'excel-sparklines.xlsx')),
	);

const extOf = (xml: string) => xml.slice(xml.indexOf('<extLst'), xml.indexOf('</extLst>') + 9);

describe('sparkline groups', () => {
	it('reads the groups Excel wrote with their type, flags, axes and colours', async () => {
		const groups = sheetSparklineGroups((await loadXlsx(fixture())).sheets[0]!);
		const byHost = new Map(groups.map((g) => [g.sparklines[0]!.host.start.row, g]));
		expect(groups).toHaveLength(4);
		const line = byHost.get(0)!;
		expect(line).toMatchObject({
			type: 'line',
			lineWeight: 1.5,
			markers: true,
			high: true,
			low: true,
			negative: true,
			first: false,
			displayEmptyCellsAs: 'gap',
			minAxisType: 'individual',
			colorSeries: { rgb: 'FF376092' },
			colorNegative: { rgb: 'FFD00000' },
		});
		expect(line.sparklines).toEqual([{ formula: 'Sheet1!A1:E1', host: parseRange('F1') }]);
		expect(byHost.get(1)).toMatchObject({ type: 'column', first: true, last: true });
		expect(byHost.get(2)).toMatchObject({ type: 'stacked', negative: true, lineWeight: 0.75 });
		expect(byHost.get(3)).toMatchObject({
			minAxisType: 'custom',
			maxAxisType: 'custom',
			manualMin: 0,
			manualMax: 10,
			displayXAxis: true,
		});
	});

	it('saves the sparkline extension of an untouched workbook unchanged and reads it back', async () => {
		const source = await JSZip.loadAsync(fixture());
		// The writer declares the revision namespace on each group instead of the root.
		const plain = (xml: string) => xml.replace(/ xmlns:xr2="[^"]*"| xr2:uid="[^"]*"/g, '');
		const before = plain(await source.file('xl/worksheets/sheet1.xml')!.async('string'));
		const bytes = await saveXlsx(await loadXlsx(fixture()));
		const saved = await (
			await JSZip.loadAsync(bytes)
		)
			.file('xl/worksheets/sheet1.xml')!
			.async('string');
		const groups = (xml: string) =>
			xml.match(/<x14:sparklineGroups[\s\S]*<\/x14:sparklineGroups>/)?.[0];
		expect(groups(plain(saved))).toBe(groups(before));
		expect(saved).toContain('xr2:uid=');
		expect(extOf(saved)).toContain('{05C60535-1F16-4fd2-B633-F4F36F0B64E0}');
		const reread = sheetSparklineGroups((await loadXlsx(bytes)).sheets[0]!);
		expect(reread).toEqual(sheetSparklineGroups((await loadXlsx(fixture())).sheets[0]!));
	});

	it('follows row inserts that rewrite the preserved extension', async () => {
		const wb = await loadXlsx(fixture());
		const sheet = wb.sheets[0]!;
		expect(sheetSparklineGroups(sheet).map((g) => g.sparklines[0]!.host.start.row)).toContain(0);
		createEditSession(wb, { recalc: false }).insertRows(0, 0, 1);
		const lines = sheetSparklineGroups(sheet).map((g) => g.sparklines[0]!);
		expect(lines.map((l) => l.host.start.row).sort()).toEqual([1, 2, 3, 4]);
		expect(lines.map((l) => l.formula)).toContain('Sheet1!A2:E2');
	});

	it('applies schema defaults and reads theme colours, date ranges and several sparklines', () => {
		const xml =
			`<extLst xmlns="${MAIN}"><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="${X14}">` +
			`<x14:sparklineGroups xmlns:xm="${XM}"><x14:sparklineGroup dateAxis="1" type="bogus" rightToLeft="1" displayHidden="1" minAxisType="group">` +
			`<x14:colorSeries theme="4" tint="-0.25"/><xm:f>Sheet1!A10:E10</xm:f><x14:sparklines>` +
			`<x14:sparkline><xm:f>Sheet1!A1:E1</xm:f><xm:sqref>F1</xm:sqref></x14:sparkline>` +
			`<x14:sparkline><xm:sqref>F2</xm:sqref></x14:sparkline><x14:sparkline><xm:f>A3:E3</xm:f></x14:sparkline>` +
			`</x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst>`;
		const [group] = readSparklineGroups([xml]);
		expect(group).toMatchObject({
			type: 'line',
			lineWeight: 0.75,
			dateAxis: true,
			dateFormula: 'Sheet1!A10:E10',
			displayEmptyCellsAs: 'zero',
			rightToLeft: true,
			displayHidden: true,
			minAxisType: 'group',
			maxAxisType: 'individual',
			colorSeries: { theme: 4, tint: -0.25 },
		});
		expect(group!.sparklines).toEqual([
			{ formula: 'Sheet1!A1:E1', host: parseRange('F1') },
			{ host: parseRange('F2') },
		]);
		expect(readSparklineGroups(['<extLst/>', 'not xml sparklineGroup'])).toEqual([]);
	});

	it('caches per sheet until the extension list changes', () => {
		const sheet = createWorkbook().sheets[0]!;
		const empty = sheetSparklineGroups(sheet);
		expect(empty).toEqual([]);
		expect(sheetSparklineGroups(sheet)).toBe(empty);
		sheet.preserved.set('extLst', [
			`<extLst xmlns="${MAIN}"><ext xmlns:x14="${X14}"><x14:sparklineGroups xmlns:xm="${XM}"><x14:sparklineGroup><x14:sparklines><x14:sparkline><xm:f>A1:B1</xm:f><xm:sqref>C1</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst>`,
		]);
		expect(sheetSparklineGroups(sheet)).toHaveLength(1);
	});
});
