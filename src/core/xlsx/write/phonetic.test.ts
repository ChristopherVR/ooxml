import { describe, expect, it } from 'vitest';
import { getCell } from '../cells';
import { createEditSession } from '../edit/session';
import { loadXlsx } from '../read/index';
import { HEAD, R, X, miniPackage, roundTrip, ws } from './mini-package-fixtures';

/** Two Japanese strings: the first has a phonetic guide with settings, the second only runs. */
const SST = `${HEAD}<sst xmlns="${X}" count="2" uniqueCount="2"><si><t>漢字</t><rPh sb="0" eb="2"><t>かんじ</t></rPh><phoneticPr fontId="0" type="fullwidthKatakana" alignment="left"/></si><si><t>東京</t><rPh sb="0" eb="2"><t>トウキョウ</t></rPh></si></sst>`;

const bytes = () =>
	miniPackage({
		sheets: [
			{
				name: 'Sheet1',
				xml: ws('<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>'),
			},
		],
		contentTypes:
			'<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>',
		workbookRels: `<Relationship Id="rIdSst" Type="${R}/sharedStrings" Target="sharedStrings.xml"/>`,
		parts: { 'xl/sharedStrings.xml': SST },
	});

describe('phonetic guides (rPh and phoneticPr)', () => {
	it('reads the guide of a shared string onto the cell', async () => {
		const wb = await loadXlsx(await bytes());
		const sheet = wb.sheets[0]!;
		expect(getCell(sheet, 0, 0)).toMatchObject({
			value: '漢字',
			phonetic: {
				runs: [{ text: 'かんじ', sb: 0, eb: 2 }],
				properties: { fontId: 0, type: 'fullwidthKatakana', alignment: 'left' },
			},
		});
		expect(getCell(sheet, 0, 1)?.phonetic).toEqual({
			runs: [{ text: 'トウキョウ', sb: 0, eb: 2 }],
		});
	});

	it('writes the guide back after the runs, with the settings, and reloads it', async () => {
		const { wb, again, part } = await roundTrip(await bytes());
		const sst = await part('xl/sharedStrings.xml');
		expect(sst).toContain(
			'<si><t>漢字</t><rPh sb="0" eb="2"><t>かんじ</t></rPh><phoneticPr fontId="0" type="fullwidthKatakana" alignment="left"/></si>',
		);
		expect(sst).toContain('<si><t>東京</t><rPh sb="0" eb="2"><t>トウキョウ</t></rPh></si>');
		expect(getCell(again.sheets[0]!, 0, 0)?.phonetic).toEqual(
			getCell(wb.sheets[0]!, 0, 0)?.phonetic,
		);
		expect(getCell(again.sheets[0]!, 0, 1)?.phonetic).toEqual(
			getCell(wb.sheets[0]!, 0, 1)?.phonetic,
		);
	});

	it('drops the guide when the cell is given a typed value', async () => {
		const wb = await loadXlsx(await bytes());
		const session = createEditSession(wb);
		session.setCellValue(0, 0, 0, '新宿');
		expect(getCell(wb.sheets[0]!, 0, 0)).toMatchObject({ value: '新宿' });
		expect(getCell(wb.sheets[0]!, 0, 0)?.phonetic).toBeUndefined();
	});
});
