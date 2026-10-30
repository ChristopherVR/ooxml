import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { putCell } from '../cells';
import { createEditSession } from '../edit/session';
import { loadXlsx } from '../read/index';
import { createWorkbook } from '../workbook';
import { saveXlsx } from './index';
import { HEAD, R, X, miniPackage, ws } from './mini-package-fixtures';

/** A package whose formula has a cached result and whose calcPr carries source attributes. */
const SOURCE_CALC_PR = '<calcPr calcId="124519" concurrentCalc="0"/>';
const bytes = () =>
	miniPackage({
		sheets: [
			{
				name: 'Sheet1',
				xml: ws('<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>A1*10</f><v>20</v></c></row>'),
			},
		],
		parts: {
			'xl/workbook.xml': `${HEAD}<workbook xmlns="${X}" xmlns:r="${R}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>${SOURCE_CALC_PR}</workbook>`,
		},
	});

/** The attributes of the saved `calcPr` element. */
async function savedCalcPr(saved: Uint8Array): Promise<Record<string, string>> {
	const xml = await JSZip.loadAsync(saved).then((zip) =>
		zip.file('xl/workbook.xml')!.async('string'),
	);
	const tag = /<calcPr\b([^>]*)\/>/.exec(xml)?.[1] ?? '';
	return Object.fromEntries([...tag.matchAll(/(\w+)="([^"]*)"/g)].map((m) => [m[1]!, m[2]!]));
}

describe('calcPr on save', () => {
	it('keeps the loaded calcPr when nothing was edited', async () => {
		const saved = await saveXlsx(await loadXlsx(await bytes()));
		expect(await savedCalcPr(saved)).toEqual({ calcId: '124519', concurrentCalc: '0' });
	});

	it('sets fullCalcOnLoad when an edited formula was not recalculated', async () => {
		const wb = await loadXlsx(await bytes());
		createEditSession(wb, { recalc: false }).setCellInput(0, 0, 1, '=A1*3');
		expect(wb.valuesStale).toBe(true);
		const attrs = await savedCalcPr(await saveXlsx(wb));
		expect(attrs).toEqual({ calcId: '124519', concurrentCalc: '0', fullCalcOnLoad: '1' });
	});

	it('sets fullCalcOnLoad for a formula written without a cached result', async () => {
		const wb = createWorkbook({ sheets: ['Sheet1'] });
		putCell(wb.sheets[0]!, 0, 0, { value: null, formula: 'NOW()' });
		expect(await savedCalcPr(await saveXlsx(wb))).toMatchObject({
			calcId: '191029',
			fullCalcOnLoad: '1',
		});
	});

	it('does not set fullCalcOnLoad after an edit the session recalculated', async () => {
		const wb = await loadXlsx(await bytes());
		createEditSession(wb).setCellInput(0, 0, 1, '=A1*3');
		expect(wb.valuesStale).toBeUndefined();
		expect(await savedCalcPr(await saveXlsx(wb))).toEqual({
			calcId: '124519',
			concurrentCalc: '0',
		});
	});

	it('keeps the workbook stale in manual calculation until it is calculated', async () => {
		const wb = await loadXlsx(await bytes());
		const session = createEditSession(wb);
		session.setCalcMode('manual');
		session.setCellValue(0, 0, 0, 5);
		expect((await savedCalcPr(await saveXlsx(wb)))['fullCalcOnLoad']).toBe('1');
		session.calculateNow();
		expect((await savedCalcPr(await saveXlsx(wb)))['fullCalcOnLoad']).toBeUndefined();
	});
});
