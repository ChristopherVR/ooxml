import { describe, expect, it } from 'vitest';
import { analyzeVisioFormula, evaluateVisioFormula, type VisioFormulaValue } from './formula';

const cells: Record<string, VisioFormulaValue> = {
	width: { value: 3, unit: 'length' },
	pinx: { value: 5, unit: 'length' },
	piny: { value: 4, unit: 'length' },
	locpinx: { value: 1, unit: 'length' },
	locpiny: { value: 1, unit: 'length' },
	angle: { value: Math.PI / 2, unit: 'angle' },
	'user.target': { value: 7, unit: 'scalar' },
	'user.name': { value: 7, unit: 'scalar' },
};
const resolve = (reference: { shapeId?: string; cell: string }): VisioFormulaValue => {
	const value =
		cells[
			`${reference.shapeId ? `sheet${reference.shapeId}!` : ''}${reference.cell.toLowerCase()}`
		];
	if (!value) throw new Error(`Unknown cell ${reference.cell}`);
	return value;
};
const run = (source: string, table = resolve) => evaluateVisioFormula(source, table);
const number = (source: string) => run(source).value;

describe('STRSAMEEX', () => {
	it('compares with Windows NORM flags', () => {
		expect(number('STRSAMEEX("abc","ABC",0)')).toBe(0);
		expect(number('STRSAMEEX("abc","ABC",1)')).toBe(1);
		expect(number('STRSAMEEX("cafe","café",2)')).toBe(1);
		expect(number('STRSAMEEX("cafe","café",0)')).toBe(0);
		expect(number('STRSAMEEX("a-b c","abc",4)')).toBe(1);
		expect(number('STRSAMEEX("A-B","ab",5)')).toBe(1);
	});
	it('refuses unrecorded flags, a locale and bad arguments', () => {
		expect(() => run('STRSAMEEX("a","a",65536)')).toThrow(/kana or width/);
		expect(() => run('STRSAMEEX("a","a",1,1033)')).toThrow(/locale/);
		expect(() => run('STRSAMEEX("a","a",-1)')).toThrow(/whole number/);
		expect(() => run('STRSAMEEX("a","a",1.5)')).toThrow(/whole number/);
		expect(() => run('STRSAMEEX("a","a")')).toThrow(/argument count/);
		expect(analyzeVisioFormula('STRSAMEEX("a","a",1)').unsupportedFunctions).toEqual([]);
	});
});

describe('FORMAT', () => {
	const len = (picture: string, value: string) => `LEN(FORMAT(${value},"${picture}"))`;
	it('applies numeric pictures', () => {
		expect(number(len('0.00', '3.14159'))).toBe(4);
		expect(number(len('0.00', '2'))).toBe(4);
		expect(number(len('0.##', '2'))).toBe(1);
		expect(number(len('0.##', '2.5'))).toBe(3);
		expect(number(len('#,##0', '1234567'))).toBe(9);
		expect(number(len('0%', '0.256'))).toBe(3);
		expect(number(len('000', '7'))).toBe(3);
		expect(number(len('0.00', '-1.005'))).toBe(5);
		expect(number('STRSAME(FORMAT(1234.5,"#,##0.00"),"1,234.50")')).toBe(1);
		expect(number('STRSAME(FORMAT(0.256,"0.0%"),"25.6%")')).toBe(1);
		expect(number('STRSAME(FORMAT(-2.5,"0"),"-3")')).toBe(1);
	});
	it('refuses pictures and values outside the subset', () => {
		for (const picture of ['yyyy', '0.00E+00', '0;-0', 'x0', '#.##', ',0', '0,', '0#'])
			expect(() => run(`LEN(FORMAT(1,"${picture}"))`), picture).toThrow(/FORMAT picture/);
		expect(() => run('LEN(FORMAT(Width,"0"))')).toThrow(/plain number/);
		expect(() => run('LEN(FORMAT(1))')).toThrow(/argument count/);
		expect(() => run('FORMAT(1,"0")')).toThrow(/does not return a number/);
	});
});

describe('DATE', () => {
	it('returns an OLE date serial that cannot pass as a plain number', () => {
		expect(run('DATE(1899,12,30)')).toEqual({ value: 0, unit: 'time' });
		expect(run('DATE(2000,1,1)')).toEqual({ value: 36526, unit: 'time' });
		expect(run('DATE(2000,3,1)-DATE(2000,2,1)')).toEqual({ value: 29, unit: 'time' });
		expect(() => run('DATE(2000,3,1)+1')).toThrow(/Incompatible/);
	});
	it('rejects impossible days', () => {
		for (const call of ['DATE(2001,2,29)', 'DATE(2000,13,1)', 'DATE(2000,0,1)', 'DATE(0,1,1)'])
			expect(() => run(call), call).toThrow();
		expect(() => run('DATE(2000,1.5,1)')).toThrow(/whole/);
		expect(() => run('DATE(2000,1)')).toThrow(/argument count/);
		expect(() => run('DATE(1 in,1,1)')).toThrow(/plain/);
		expect(number('IFERROR(DATE(2001,2,29),-1)')).toBe(-1);
	});
});

describe('ANGLETOLOC and LOCTOPAR', () => {
	it('converts an angle and a point between parent and local frames', () => {
		const angle = run('ANGLETOLOC(SHEETREF(),180 deg)');
		expect(angle.unit).toBe('angle');
		expect(angle.value).toBeCloseTo(Math.PI / 2, 12);
		// Local (2,1) is one inch right of the pin; a quarter turn puts it one inch above it.
		expect(run('PNTX(LOCTOPAR(SHEETREF(),2,1))').value).toBeCloseTo(5, 12);
		expect(run('PNTY(LOCTOPAR(SHEETREF(),2,1))').value).toBeCloseTo(5, 12);
		expect(run('PNTX(LOCTOPAR(SHEETREF(),1,2))').value).toBeCloseTo(4, 12);
	});
	it('refuses other sheets, flips and misuse', () => {
		expect(() => run('ANGLETOLOC(1,0)')).toThrow(/sheet other than/);
		expect(() => run('LOCTOPAR(SHEETREF(),1,1)')).toThrow(/does not return a number/);
		expect(() => run('PNTX(LOCTOPAR(SHEETREF(),1))')).toThrow(/argument count/);
		expect(() => run('ANGLETOLOC(SHEETREF(),1 in)')).toThrow(/angle/);
		const flipped = (reference: { cell: string }) =>
			reference.cell === 'FlipX' ? { value: 1, unit: 'scalar' as const } : resolve(reference);
		expect(() => run('ANGLETOLOC(SHEETREF(),0)', flipped)).toThrow(/flipped/);
	});
	it('records the frame cells as dependencies', () => {
		const cellsOf = (source: string) =>
			analyzeVisioFormula(source).references.map((reference) => reference.cell);
		expect(cellsOf('ANGLETOLOC(SHEETREF(),0)')).toEqual(['Angle', 'FlipX', 'FlipY']);
		expect(cellsOf('PNTX(LOCTOPAR(SHEETREF(),Width,1))')).toEqual(
			expect.arrayContaining(['PinX', 'PinY', 'LocPinX', 'LocPinY', 'Angle', 'Width']),
		);
		expect(analyzeVisioFormula('ANGLETOLOC(SHEETREF(),0)').unsupportedFunctions).toEqual([]);
	});
});

describe('INDIRECT', () => {
	it('reads the cell a text address names', () => {
		expect(number('INDIRECT("User.Target")')).toBe(7);
		expect(number('INDIRECT("User."&"Name")*2')).toBe(14);
		expect(
			run('INDIRECT("Sheet.2!Width")', (ref) =>
				resolve({ cell: ref.shapeId === '2' ? 'Width' : 'x' }),
			).unit,
		).toBe('length');
	});
	it('fails on an address that is not a cell', () => {
		expect(() => run('INDIRECT("1+")')).toThrow();
		expect(() => run('INDIRECT("Width+1")')).toThrow(/cell address/);
		expect(() => run('INDIRECT("Missing")')).toThrow(/Unknown cell/);
		expect(() => run('INDIRECT()')).toThrow(/argument count/);
	});
	it('is static for a literal address only when asked, and dynamic for a computed one', () => {
		const literal = analyzeVisioFormula('INDIRECT("User.Target")', { literalIndirect: true });
		expect(literal).toMatchObject({ dynamic: false, unsupportedFunctions: [] });
		expect(literal.references).toEqual([{ cell: 'User.Target' }]);
		expect(
			analyzeVisioFormula('INDIRECT("Sheet.3!Width")', { literalIndirect: true }).references,
		).toEqual([{ shapeId: '3', cell: 'Width' }]);
		expect(analyzeVisioFormula('INDIRECT("User.Target")').dynamic).toBe(true);
		expect(analyzeVisioFormula('INDIRECT("User."&"Name")', { literalIndirect: true }).dynamic).toBe(
			true,
		);
		expect(analyzeVisioFormula('INDIRECT(Width)').dynamic).toBe(true);
		expect(analyzeVisioFormula('INDIRECT("1+")').dynamic).toBe(true);
	});
});
