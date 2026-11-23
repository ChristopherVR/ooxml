import { describe, expect, it } from 'vitest';
import { createWorkbook } from '../../workbook';
import { createCalcEngine } from '../engine';
import { calc, E } from '../test-helpers';
import { FUNCTION_CATALOG } from './registry';

describe('INFO', () => {
	it('returns workbook and environment metadata for each documented type_text', () => {
		const wb = createWorkbook({ sheets: ['Sheet1', 'Sheet2'] });
		const sheet = wb.sheets[0]!;
		sheet.view.topLeft = { row: 2, col: 3 };
		sheet.rows.set(
			0,
			new Map([
				[0, { value: null, formula: 'INFO("directory")' }],
				[1, { value: null, formula: 'INFO("numfile")' }],
				[2, { value: null, formula: 'INFO("origin")' }],
				[3, { value: null, formula: 'INFO("osversion")' }],
				[4, { value: null, formula: 'INFO("recalc")' }],
				[5, { value: null, formula: 'INFO("release")' }],
				[6, { value: null, formula: 'INFO("SYSTEM")' }],
				[7, { value: null, formula: 'INFO("unknown")' }],
			]),
		);
		createCalcEngine(wb).recalculateAll();
		const row = sheet.rows.get(0)!;
		expect(row.get(0)?.value).toBe('');
		expect(row.get(1)?.value).toBe(2);
		expect(row.get(2)?.value).toBe('$A:$D$3');
		expect(row.get(3)?.value).toBe('Web');
		expect(row.get(4)?.value).toBe('Automatic');
		expect(row.get(5)?.value).toBe('16.0');
		expect(row.get(6)?.value).toBe('pcdos');
		expect(row.get(7)?.value).toEqual(E.VALUE);

		wb.calcMode = 'manual';
		createCalcEngine(wb).recalculateAll();
		expect(sheet.rows.get(0)?.get(4)?.value).toBe('Manual');
	});

	it('propagates errors and rejects blank type_text', () => {
		expect(calc('INFO(1/0)')).toEqual(E.DIV0);
		expect(calc('INFO(A1)')).toEqual(E.VALUE);
	});
});

describe('ASC, DBCS, and JIS', () => {
	it('converts full-width ASCII, spaces, and Katakana (including dakuten/handakuten) to half-width', () => {
		expect(calc('ASC("ＥＸＣＥＬ　１２３！")')).toBe('EXCEL 123!');
		expect(calc('ASC("カタカナ・ガギグゲゴ・パピプペポ・ヴ")')).toBe('ｶﾀｶﾅ･ｶﾞｷﾞｸﾞｹﾞｺﾞ･ﾊﾟﾋﾟﾌﾟﾍﾟﾎﾟ･ｳﾞ');
		expect(calc('ASC("「東京」、テスト。")')).toBe('｢東京｣､ﾃｽﾄ｡');
	});

	it('converts half-width ASCII, spaces, and Katakana to full-width via JIS and DBCS', () => {
		expect(calc('JIS("EXCEL 123!")')).toBe('ＥＸＣＥＬ　１２３！');
		expect(calc('DBCS("EXCEL 123!")')).toBe('ＥＸＣＥＬ　１２３！');
		expect(calc('JIS("ｶﾀｶﾅ･ｶﾞｷﾞｸﾞｹﾞｺﾞ･ﾊﾟﾋﾟﾌﾟﾍﾟﾎﾟ･ｳﾞ")')).toBe('カタカナ・ガギグゲゴ・パピプペポ・ヴ');
		expect(calc('DBCS("｢東京｣､ﾃｽﾄ｡")')).toBe('「東京」、テスト。');
	});

	it('round-trips ASC and JIS/DBCS and propagates errors', () => {
		expect(calc('ASC(JIS("Hello ｶﾞﾊﾟ 123"))')).toBe('Hello ｶﾞﾊﾟ 123');
		expect(calc('DBCS(ASC("Ｈｅｌｌｏ　ガパ　１２３"))')).toBe('Ｈｅｌｌｏ　ガパ　１２３');
		expect(calc('ASC(NA())')).toEqual(E.NA);
		expect(calc('DBCS(1/0)')).toEqual(E.DIV0);
		expect(calc('JIS(1/0)')).toEqual(E.DIV0);
	});
});

describe('PHONETIC', () => {
	it('extracts phonetic runs from a cell or upper-left cell of a range', () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0]!;
		sheet.rows.set(
			0,
			new Map([
				[
					0,
					{
						value: '東京都',
						phonetic: { runs: [{ text: 'トウキョウト', sb: 0, eb: 3 }] },
					},
				],
				[
					1,
					{
						value: '東京大阪',
						phonetic: {
							runs: [
								{ text: 'トウキョウ', sb: 0, eb: 2 },
								{ text: 'オオサカ', sb: 2, eb: 4 },
							],
						},
					},
				],
				[2, { value: 'plain text' }],
				[3, { value: 42 }],
				[4, { value: true }],
			]),
		);
		sheet.rows.set(
			1,
			new Map([
				[0, { value: null, formula: 'PHONETIC(A1)' }],
				[1, { value: null, formula: 'PHONETIC(B1)' }],
				[2, { value: null, formula: 'PHONETIC(A1:C1)' }],
				[3, { value: null, formula: 'PHONETIC(C1)' }],
				[4, { value: null, formula: 'PHONETIC(D1)' }],
				[5, { value: null, formula: 'PHONETIC(E1)' }],
				[6, { value: null, formula: 'PHONETIC(F1)' }],
			]),
		);
		createCalcEngine(wb).recalculateAll();
		const results = sheet.rows.get(1)!;
		expect(results.get(0)?.value).toBe('トウキョウト');
		expect(results.get(1)?.value).toBe('トウキョウオオサカ');
		expect(results.get(2)?.value).toBe('トウキョウト');
		expect(results.get(3)?.value).toBe('plain text');
		expect(results.get(4)?.value).toBe('');
		expect(results.get(5)?.value).toBe('');
		expect(results.get(6)?.value).toBe('');
	});

	it('returns #VALUE! for non-references and #N/A for multi-area unions', () => {
		expect(calc('PHONETIC("abc")')).toEqual(E.VALUE);
		expect(calc('PHONETIC((A1,B1))', { A1: 'a', B1: 'b' })).toEqual(E.NA);
		expect(calc('PHONETIC(A1)', { A1: E.DIV0 })).toEqual(E.DIV0);
	});
});

describe('EUROCONVERT', () => {
	it('matches Microsoft Excel examples for EUR, member currencies, and triangulation', () => {
		expect(calc('EUROCONVERT(1.2,"DEM","EUR")')).toBe(0.61);
		expect(calc('EUROCONVERT(1,"FRF","EUR",TRUE)') as number).toBeCloseTo(0.152449017, 8);
		expect(calc('EUROCONVERT(1,"FRF","DEM",TRUE,3)') as number).toBeCloseTo(0.29728616, 8);
		expect(calc('EUROCONVERT(1,"FRF","DEM",FALSE,3)')).toBe(0.3);
		expect(calc('EUROCONVERT(1,"EUR","ITL")')).toBe(1936);
		expect(calc('EUROCONVERT(100,"EUR","HRK")')).toBe(753.45);
		expect(calc('EUROCONVERT(1.234,"EUR","EUR")')).toBe(1.234);
		expect(calc('EUROCONVERT(1,"dem","eur")')).toBe(0.51);
	});

	it('rejects unknown currency codes and triangulation precision below 3', () => {
		expect(calc('EUROCONVERT(10,"USD","EUR")')).toEqual(E.VALUE);
		expect(calc('EUROCONVERT(10,"EUR","GBP")')).toEqual(E.VALUE);
		expect(calc('EUROCONVERT(10,"FRF","DEM",FALSE,2)')).toEqual(E.VALUE);
		expect(calc('EUROCONVERT(1/0,"DEM","EUR")')).toEqual(E.DIV0);
	});
});

describe('FUNCTION_CATALOG', () => {
	it('lists INFO, ASC, DBCS, JIS, PHONETIC, and EUROCONVERT', () => {
		const names = new Set(FUNCTION_CATALOG.map((f) => f.name));
		for (const name of ['INFO', 'ASC', 'DBCS', 'JIS', 'PHONETIC', 'EUROCONVERT'])
			expect(names.has(name)).toBe(true);
	});
});
