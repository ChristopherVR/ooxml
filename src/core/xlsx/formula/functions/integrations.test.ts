import { describe, expect, it, vi } from 'vitest';
import { putCell } from '../../cells';
import { createMemoryCubeProvider } from '../cube-memory';
import { createCalcEngine, type CalcExternalProviders } from '../engine';
import { ERR } from '../values';
import { book, E, get, grid, locate, set } from '../test-helpers';

const UNCONFIGURED_CASES: readonly [string, string][] = [
	['CUBEKPIMEMBER', 'CUBEKPIMEMBER("Sales","Margin",1)'],
	['CUBEMEMBER', 'CUBEMEMBER("Sales","[Product].[Bikes]")'],
	['CUBEMEMBERPROPERTY', 'CUBEMEMBERPROPERTY("Sales","[Product].[Bikes]","Color")'],
	['CUBERANKEDMEMBER', 'CUBERANKEDMEMBER("Sales","[Product].Members",1)'],
	['CUBESET', 'CUBESET("Sales","[Product].Members")'],
	['CUBESETCOUNT', 'CUBESETCOUNT(A1)'],
	['CUBEVALUE', 'CUBEVALUE("Sales","[Measures].[Revenue]")'],
	['COPILOT', 'COPILOT("Summarize",A1:A2)'],
	['STOCKHISTORY', 'STOCKHISTORY("MSFT",45292,45295)'],
	['TRANSLATE', 'TRANSLATE("hello","en","fr")'],
	['DETECTLANGUAGE', 'DETECTLANGUAGE("bonjour")'],
	['IMAGE', 'IMAGE("https://example.test/logo.png","Logo")'],
	['RTD', 'RTD("my.rtd","","MSFT","LAST")'],
	['WEBSERVICE', 'WEBSERVICE("https://example.test/data")'],
];

describe('external integration functions without providers', () => {
	it.each(UNCONFIGURED_CASES)('%s returns #NAME? when newly entered', (_name, formula) => {
		const wb = book({ A1: 1, A2: 2 });
		const calc = createCalcEngine(wb);
		calc.recalculateAll();
		set(wb, 'B1', `=${formula}`);
		calc.recalculateFrom([locate(wb, 'B1')]);
		expect(get(wb, 'B1')).toEqual(E.NAME);
	});

	it.each(UNCONFIGURED_CASES)('%s preserves cached values on initial load', (_name, formula) => {
		const wb = book({ A1: 1, A2: 2 });
		const sheet = wb.sheets[0]!;
		putCell(sheet, 0, 1, { value: 'cached-ok', formula });
		const calc = createCalcEngine(wb);
		calc.recalculateAll();
		expect(get(wb, 'B1')).toBe('cached-ok');
	});
});

describe('CUBE* functions with createMemoryCubeProvider', () => {
	const cube = createMemoryCubeProvider({
		Sales: {
			members: {
				'[Product].[Bikes]': { caption: 'Bikes', properties: { Category: 'Cycles', Margin: 0.42 } },
				'[Product].[Helmets]': { caption: 'Helmets', properties: { Category: 'Safety' } },
				'[Year].[2024]': { caption: 'FY 2024' },
				'[Measures].[Revenue]': { caption: 'Revenue' },
				'[Product].[Bikes]|[Year].[2024]': { caption: 'Bikes 2024' },
			},
			kpis: {
				'Gross Margin': {
					1: { expression: '[KPI].[Gross Margin].[Value]', caption: 'GM Value', value: 0.38 },
					2: { expression: '[KPI].[Gross Margin].[Goal]', caption: 'GM Goal', value: 0.4 },
				},
			},
			sets: {
				'[Product].Members': [
					{ expression: '[Product].[Bikes]', caption: 'Bikes', sortValue: 500 },
					{ expression: '[Product].[Helmets]', caption: 'Helmets', sortValue: 120 },
					{ expression: '[Product].[Gloves]', caption: 'Gloves', sortValue: 260 },
				],
			},
			values: {
				'[Measures].[Revenue]|[Product].[Bikes]': 500,
				'[Measures].[Revenue]|[Product].[Gloves]': 260,
				'[Measures].[Revenue]|[Product].[Helmets]': 120,
				'[Measures].[Revenue]|[Product].[Bikes]|[Year].[2024]': 480,
				'[Product].[Bikes]|[Product].[Gloves]|[Product].[Helmets]': 880,
			},
		},
	});

	it('evaluates CUBEMEMBER, CUBESET, CUBESETCOUNT, CUBERANKEDMEMBER, CUBEVALUE, CUBEKPIMEMBER and CUBEMEMBERPROPERTY end to end', () => {
		const wb = book({
			A1: '=CUBEMEMBER("Sales","[Measures].[Revenue]")',
			A2: '=CUBEMEMBER("Sales",{"[Product].[Bikes]","[Year].[2024]"})',
			A3: '=CUBESET("Sales","[Product].Members","Top Products",2,"[Measures].[Revenue]")',
			A4: '=CUBESETCOUNT(A3)',
			A5: '=CUBERANKEDMEMBER("Sales",A3,1)',
			A6: '=CUBERANKEDMEMBER("Sales",A3,2)',
			A7: '=CUBEVALUE("Sales",A1,A5)',
			A8: '=CUBEVALUE("Sales",A1,A6)',
			A9: '=CUBEVALUE("Sales",A3)',
			A10: '=CUBEMEMBERPROPERTY("Sales",A5,"Margin")',
			A11: '=CUBEKPIMEMBER("Sales","Gross Margin",1)',
			A12: '=CUBEVALUE("Sales",A11)',
			A13: '=CUBESETCOUNT(CUBESET("Sales","[Product].Members","Inline"))',
		});
		createCalcEngine(wb, { providers: { cube } }).recalculateAll();
		expect(get(wb, 'A1')).toBe('Revenue');
		expect(get(wb, 'A2')).toBe('Bikes 2024');
		expect(get(wb, 'A3')).toBe('Top Products');
		expect(get(wb, 'A4')).toBe(3);
		expect(get(wb, 'A5')).toBe('Bikes');
		expect(get(wb, 'A6')).toBe('Gloves');
		expect(get(wb, 'A7')).toBe(500);
		expect(get(wb, 'A8')).toBe(260);
		expect(get(wb, 'A9')).toBe(880);
		expect(get(wb, 'A10')).toBe(0.42);
		expect(get(wb, 'A11')).toBe('GM Value');
		expect(get(wb, 'A12')).toBe(0.38);
		expect(get(wb, 'A13')).toBe(3);
	});

	it('validates CUBE* arguments', () => {
		const wb = book({
			A1: '=CUBEKPIMEMBER("Sales","Gross Margin",7)',
			A2: '=CUBERANKEDMEMBER("Sales","[Product].Members",0)',
			A3: '=CUBESET("Sales","[Product].Members","Bad",1)',
			A4: '=CUBESET("Sales","[Product].Members","Bad",9)',
			A5: '=CUBESETCOUNT("not-a-set")',
			A6: '=CUBEMEMBER("","[Product].[Bikes]")',
			A7: '=CUBERANKEDMEMBER("Sales","[Product].Members",99)',
		});
		createCalcEngine(wb, { providers: { cube } }).recalculateAll();
		expect(get(wb, 'A1')).toEqual(E.VALUE);
		expect(get(wb, 'A2')).toEqual(E.VALUE);
		expect(get(wb, 'A3')).toEqual(E.VALUE);
		expect(get(wb, 'A4')).toEqual(E.VALUE);
		expect(get(wb, 'A5')).toEqual(E.VALUE);
		expect(get(wb, 'A6')).toEqual(E.VALUE);
		expect(get(wb, 'A7')).toEqual(E.NA);
	});
});

describe('COPILOT, STOCKHISTORY, TRANSLATE, DETECTLANGUAGE, IMAGE, RTD and WEBSERVICE', () => {
	it('evaluates COPILOT for scalar and spilled matrix responses', () => {
		const evaluate = vi.fn<NonNullable<CalcExternalProviders['copilot']>['evaluate']>((req) => {
			if (req.parts[0]?.prompt === 'table') {
				return [
					['Item', 'Score'],
					['A', 10],
				];
			}
			const ctx = req.parts[0]?.context;
			return `${req.parts[0]?.prompt}:${JSON.stringify(ctx)}`;
		});
		const wb = book({
			A1: 10,
			A2: 20,
			B1: '=COPILOT("sum",A1:A2)',
			C1: '=COPILOT("table")',
			E1: '=COPILOT("")',
		});
		createCalcEngine(wb, { providers: { copilot: { evaluate } } }).recalculateAll();
		expect(get(wb, 'B1')).toBe('sum:[[10],[20]]');
		expect(grid(wb, 'C1:D2')).toEqual([
			['Item', 'Score'],
			['A', 10],
		]);
		expect(get(wb, 'E1')).toEqual(E.VALUE);
	});

	it('evaluates STOCKHISTORY with headers, custom properties and validation', () => {
		const stockHistory: NonNullable<CalcExternalProviders['stockHistory']> = {
			history: (req) => {
				if (req.stock === 'EMPTY') return [];
				return [
					{ date: 45292, open: 100, high: 105, low: 99, close: 104, volume: 1000 },
					{ date: 45293, open: 104, high: 110, low: 103, close: 109, volume: 1500 },
				];
			},
		};
		const wb = book({
			A1: '=STOCKHISTORY("MSFT",45292,45293,0,2,0,1,5)',
			E1: '=STOCKHISTORY("MSFT",45292,45293,0,0,1)',
			G1: '=STOCKHISTORY("MSFT",45295,45292)',
			G2: '=STOCKHISTORY("MSFT",45292,45293,3)',
			G3: '=STOCKHISTORY("MSFT",45292,45293,0,1,6)',
			G4: '=STOCKHISTORY("EMPTY",45292,45293,0,0)',
		});
		createCalcEngine(wb, { providers: { stockHistory } }).recalculateAll();
		expect(grid(wb, 'A1:C4')).toEqual([
			['MSFT', '', ''],
			['Date', 'Close', 'Volume'],
			[45292, 104, 1000],
			[45293, 109, 1500],
		]);
		expect(grid(wb, 'E1:E2')).toEqual([[104], [109]]);
		expect(get(wb, 'G1')).toEqual(E.VALUE);
		expect(get(wb, 'G2')).toEqual(E.VALUE);
		expect(get(wb, 'G3')).toEqual(E.VALUE);
		expect(get(wb, 'G4')).toEqual(E.NA);
	});

	it('evaluates TRANSLATE and DETECTLANGUAGE', () => {
		const providers: CalcExternalProviders = {
			translate: {
				translate: (text, src, dst) => `${text}[${src ?? 'auto'}->${dst ?? 'en'}]`,
			},
			detectLanguage: {
				detectLanguage: (text) => (text.includes('bonjour') ? 'fr' : 'en'),
			},
		};
		const wb = book({
			A1: '=TRANSLATE("hello","en","fr")',
			A2: '=TRANSLATE("hola")',
			A3: '=TRANSLATE("hi","","fr")',
			A4: '=DETECTLANGUAGE("bonjour")',
		});
		createCalcEngine(wb, { providers }).recalculateAll();
		expect(get(wb, 'A1')).toBe('hello[en->fr]');
		expect(get(wb, 'A2')).toBe('hola[auto->en]');
		expect(get(wb, 'A3')).toEqual(E.VALUE);
		expect(get(wb, 'A4')).toBe('fr');
	});

	it('evaluates IMAGE with URL scheme and sizing validation', () => {
		const resolve = vi.fn<NonNullable<CalcExternalProviders['image']>['resolve']>(
			(req) => `${req.altText || req.source}(${req.sizing}:${req.width ?? ''}x${req.height ?? ''})`,
		);
		const wb = book({
			A1: '=IMAGE("https://example.test/a.png","Logo")',
			A2: '=IMAGE("https://example.test/b.png","Custom",3,120,240)',
			A3: '=IMAGE("javascript:alert(1)")',
			A4: '=IMAGE("file:///etc/passwd")',
			A5: '=IMAGE("https://example.test/c.png","Bad",3,120)',
			A6: '=IMAGE("https://example.test/d.png","Bad",0,120,240)',
			A7: '=IMAGE("https://example.test/e.png","Bad",3,-1,10)',
		});
		createCalcEngine(wb, { providers: { image: { resolve } } }).recalculateAll();
		expect(get(wb, 'A1')).toBe('Logo(0:x)');
		expect(get(wb, 'A2')).toBe('Custom(3:240x120)');
		expect(get(wb, 'A3')).toEqual(ERR.BLOCKED);
		expect(get(wb, 'A4')).toEqual(ERR.BLOCKED);
		expect(get(wb, 'A5')).toEqual(E.VALUE);
		expect(get(wb, 'A6')).toEqual(E.VALUE);
		expect(get(wb, 'A7')).toEqual(E.VALUE);
		expect(resolve).toHaveBeenCalledTimes(2);
	});

	it('evaluates RTD as volatile when configured', () => {
		let tick = 100;
		const providers: CalcExternalProviders = {
			rtd: {
				read: (progId, _server, topics) => `${progId}:${topics.join('/')}:${++tick}`,
			},
		};
		const wb = book({
			A1: 1,
			B1: '=RTD("quote.rtd","","MSFT","LAST")',
			C1: '=RTD("","","MSFT")',
		});
		const calc = createCalcEngine(wb, { providers });
		calc.recalculateAll();
		expect(get(wb, 'B1')).toBe('quote.rtd:MSFT/LAST:101');
		expect(get(wb, 'C1')).toEqual(E.VALUE);
		// Editing an unrelated cell triggers recalculation of volatile RTD.
		set(wb, 'A1', 2);
		calc.recalculateFrom([locate(wb, 'A1')]);
		expect(get(wb, 'B1')).toBe('quote.rtd:MSFT/LAST:102');
	});

	it('evaluates WEBSERVICE with URL and length validation', () => {
		const fetch = vi.fn<NonNullable<CalcExternalProviders['webservice']>['fetch']>((url) =>
			url.endsWith('/big') ? 'x'.repeat(32_768) : `ok:${url}`,
		);
		const wb = book({
			A1: '=WEBSERVICE("https://example.test/api")',
			A2: '=WEBSERVICE("ftp://example.test/file")',
			A3: `=WEBSERVICE("https://example.test/${'a'.repeat(2050)}")`,
			A4: '=WEBSERVICE("https://example.test/big")',
		});
		createCalcEngine(wb, { providers: { webservice: { fetch } } }).recalculateAll();
		expect(get(wb, 'A1')).toBe('ok:https://example.test/api');
		expect(get(wb, 'A2')).toEqual(E.VALUE);
		expect(get(wb, 'A3')).toEqual(E.VALUE);
		expect(get(wb, 'A4')).toEqual(E.VALUE);
		expect(fetch).toHaveBeenCalledTimes(2);
	});
});
