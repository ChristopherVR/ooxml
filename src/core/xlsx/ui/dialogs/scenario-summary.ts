// Scenario Manager > Summary: a standard "Scenario Summary" report on a new sheet before the
// scenarios' sheet. The structure, styles, outline and note text follow what Excel 16 writes
// (Scenarios.CreateSummary, report type 1): a purple title band with the scenario names, a hidden
// comment row, the changing cells (gray where a scenario sets them) and the result cells, each as
// plain values. Result values come from showing every scenario on a scratch copy of the workbook.
import {
	type Border,
	type CellAddress,
	type CellValue,
	columnLabel,
	createEditSession,
	type DefinedName,
	expandCells,
	formatAddress,
	getCell,
	parseAddress,
	parseCellInput,
	scenariosOf,
	styleAt,
	type StylePatch,
	type EditSession,
	type Workbook,
	type Worksheet,
} from '../../index';

export type ScenarioSummaryProblem = 'no-scenarios' | 'bad-result-cells';

export class ScenarioSummaryError extends Error {
	constructor(readonly problem: ScenarioSummaryProblem) {
		super(
			problem === 'no-scenarios'
				? 'There are no scenarios on this sheet to summarize.'
				: 'The result cells reference is not valid.',
		);
	}
}

export interface ScenarioSummaryRequest {
	/** The sheet whose scenarios are summarized. */
	sheet: number;
	/** The result cells, for example `B3:B4,D1`; empty leaves the result block out. */
	resultCells?: string;
}

const NOTES = [
	'Notes:  Current Values column represents values of changing cells at',
	'time Scenario Summary Report was created.  Changing cells for each',
	'scenario are highlighted in gray.',
];

const PURPLE = { indexed: 20 };
const GRAY = { indexed: 22 };
const WHITE = { indexed: 9 };
const BLACK = { indexed: 8 };
const NAVY = { indexed: 18 };
const solid = (color: { indexed: number }): NonNullable<StylePatch['fill']> => ({
	type: 'pattern',
	pattern: 'solid',
	fgColor: color,
	bgColor: { indexed: 24 },
});
const edge = (
	style: 'thin' | 'medium',
): { style: 'thin' | 'medium'; color: { indexed: number } } => ({
	style,
	color: { indexed: 64 },
});
const border = (top?: 'thin' | 'medium', bottom?: 'thin' | 'medium'): Partial<Border> => ({
	...(top ? { top: edge(top) } : {}),
	...(bottom ? { bottom: edge(bottom) } : {}),
});

/** The name of a defined name that is exactly this cell on this sheet, else the `$B$2` address. */
function cellLabel(workbook: Workbook, sheetIndex: number, address: CellAddress): string {
	const sheetName = workbook.sheets[sheetIndex]?.name ?? '';
	const target = formatAddress(address);
	const named = workbook.definedNames.find((name: DefinedName) => {
		if (name.localSheet !== undefined && name.localSheet !== sheetIndex) return false;
		const match = /^(?:'((?:[^']|'')+)'|([^'!]+))!\$?([A-Z]{1,3})\$?(\d+)$/i.exec(
			name.formula.trim(),
		);
		if (!match) return false;
		const sheet = (match[1] ?? match[2] ?? '').replace(/''/g, "'");
		return sheet === sheetName && `${match[3]!.toUpperCase()}${match[4]}` === target;
	});
	return named?.name ?? `$${columnLabel(address.col)}$${address.row + 1}`;
}

function scenarioValue(text: string, date1904: boolean): CellValue {
	const parsed = parseCellInput(text, { date1904 });
	return parsed.formula !== undefined ? text : (parsed.value as CellValue);
}

/** The result cells under each scenario, computed on a scratch copy. */
function resultsPerScenario(
	workbook: Workbook,
	sheetIndex: number,
	count: number,
	results: CellAddress[],
): CellValue[][] {
	const scratch = createEditSession({
		...workbook,
		sheets: structuredClone(workbook.sheets),
		definedNames: structuredClone(workbook.definedNames),
	});
	const sheet = scratch.workbook.sheets[sheetIndex] as Worksheet;
	const out: CellValue[][] = [];
	for (let i = 0; i < count; i++) {
		scratch.showScenario(sheetIndex, i);
		out.push(results.map((a) => getCell(sheet, a.row, a.col)?.value ?? null));
		scratch.undo();
	}
	return out;
}

/** Builds the summary sheet in one undo step; returns its index. */
export function createScenarioSummary(
	session: EditSession,
	request: ScenarioSummaryRequest,
): number {
	const { workbook } = session;
	const s = request.sheet;
	const sheet = workbook.sheets[s] as Worksheet;
	const scenarios = scenariosOf(sheet);
	if (!scenarios.length) throw new ScenarioSummaryError('no-scenarios');
	const resultText = request.resultCells?.trim() ?? '';
	const results = resultText ? expandCells(resultText) : [];
	if (!results) throw new ScenarioSummaryError('bad-result-cells');

	// Changing cells: every cell any scenario sets, in order of first appearance.
	const changing: CellAddress[] = [];
	const seen = new Set<string>();
	for (const scenario of scenarios)
		for (const cell of scenario.cells) {
			const address = parseAddress(cell.ref);
			if (!address || cell.deleted || seen.has(formatAddress(address))) continue;
			seen.add(formatAddress(address));
			changing.push(address);
		}
	const current = (a: CellAddress): CellValue => getCell(sheet, a.row, a.col)?.value ?? null;
	const resultValues = results.length
		? resultsPerScenario(workbook, s, scenarios.length, results)
		: [];

	const last = 4 + scenarios.length; // 0-based last column (D = 3 holds the current values)
	const lastCol = last;
	const rows: CellValue[][] = [];
	const put = (row: number, col: number, value: CellValue): void => {
		const line = (rows[row] ??= []);
		line[col] = value;
	};
	put(1, 1, 'Scenario Summary');
	put(2, 3, 'Current Values:');
	scenarios.forEach((scenario, k) => {
		put(2, 4 + k, scenario.name);
		put(3, 4 + k, scenario.comment ?? null);
	});
	put(4, 1, 'Changing Cells:');
	const changingStart = 5;
	changing.forEach((address, i) => {
		const row = changingStart + i;
		put(row, 2, cellLabel(workbook, s, address));
		put(row, 3, current(address));
		scenarios.forEach((scenario, k) => {
			const own = scenario.cells.find((c) => c.ref === formatAddress(address) && !c.deleted);
			put(row, 4 + k, own ? scenarioValue(own.value, workbook.date1904) : current(address));
		});
	});
	let next = changingStart + changing.length;
	const resultHeader = next;
	const resultStart = next + 1;
	if (results.length) {
		put(resultHeader, 1, 'Result Cells:');
		results.forEach((address, i) => {
			const row = resultStart + i;
			put(row, 2, cellLabel(workbook, s, address));
			put(row, 3, current(address));
			scenarios.forEach((_scenario, k) => put(row, 4 + k, resultValues[k]?.[i] ?? null));
		});
		next = resultStart + results.length;
	}
	NOTES.forEach((note, i) => put(next + i, 1, note));
	const resultEnd = resultStart + results.length - 1;

	let created = -1;
	session.batch('Scenario summary', () => {
		const taken = new Set(workbook.sheets.map((w) => w.name.toLowerCase()));
		let name = 'Scenario Summary';
		for (let n = 2; taken.has(name.toLowerCase()); n++) name = `Scenario Summary ${n}`;
		created = session.addSheet(name, s);
		const t = created;
		const width = rows.reduce((max, line) => Math.max(max, line.length), 0);
		session.setRangeValues(
			t,
			{ row: 0, col: 0 },
			Array.from({ length: rows.length }, (_v, r) =>
				Array.from({ length: width }, (_w, c) => rows[r]?.[c] ?? null),
			),
		);
		const range = (r0: number, c0: number, r1: number, c1: number) => [
			{ start: { row: r0, col: c0 }, end: { row: r1, col: c1 } },
		];
		const style = (r0: number, c0: number, r1: number, c1: number, patch: StylePatch) =>
			session.applyStyle(t, range(r0, c0, r1, c1), patch);
		const bold = (size: number, color: { indexed: number }) => ({ bold: true, size, color });
		// Title band.
		style(1, 1, 1, 2, {
			font: bold(12, WHITE),
			fill: solid(PURPLE),
			border: border('medium'),
			alignment: { horizontal: 'left' },
		});
		style(1, 3, 1, lastCol, {
			font: { size: 10, color: WHITE },
			fill: solid(PURPLE),
			border: border('medium'),
			alignment: { horizontal: 'right' },
		});
		style(2, 1, 2, 2, {
			font: bold(12, WHITE),
			fill: solid(PURPLE),
			border: border(undefined, 'thin'),
			alignment: { horizontal: 'left' },
		});
		style(2, 3, 2, lastCol, {
			font: { size: 10, color: WHITE },
			fill: solid(PURPLE),
			border: border(undefined, 'thin'),
			alignment: { horizontal: 'right' },
		});
		// Comment row.
		const label: StylePatch = {
			font: bold(11, BLACK),
			fill: solid(GRAY),
			alignment: { horizontal: 'left' },
		};
		style(3, 1, 3, 2, label);
		style(3, 4, 3, lastCol, {
			font: { size: 8 },
			alignment: { vertical: 'top', wrapText: true },
		});
		// Section headers.
		const sections = results.length ? [4, resultHeader] : [4];
		for (const row of sections) {
			style(row, 1, row, 2, {
				font: bold(11, NAVY),
				fill: solid(GRAY),
				border: border('thin', 'thin'),
				alignment: { horizontal: 'left' },
			});
			style(row, 3, row, lastCol, { border: border('thin', 'thin') });
		}
		for (let i = 0; i < changing.length; i++) {
			const row = changingStart + i;
			style(row, 1, row, 2, label);
			style(row, 4, row, lastCol, { fill: solid(GRAY) });
		}
		for (let i = 0; i < results.length; i++) {
			const row = resultStart + i;
			const lastResult = row === resultEnd;
			style(row, 1, row, 2, {
				...label,
				...(lastResult ? { border: border(undefined, 'medium') } : {}),
			});
			if (lastResult) style(row, 3, row, lastCol, { border: border(undefined, 'medium') });
		}
		// Values keep the number format of the cell they came from.
		const copyFormat = (row: number, address: CellAddress, firstCol: number) => {
			const numFmt = styleAt(workbook, getCell(sheet, address.row, address.col)?.styleId).numFmt;
			if (numFmt && numFmt !== 'General') style(row, firstCol, row, lastCol, { numFmt });
		};
		changing.forEach((address, i) => copyFormat(changingStart + i, address, 3));
		results.forEach((address, i) => copyFormat(resultStart + i, address, 3));
		// Columns, outline and view.
		session.setColumnWidth(t, [2], 5.09);
		const valueCols = Array.from({ length: lastCol - 2 }, (_v, i) => 3 + i);
		session.setColumnWidth(t, valueCols, 12.18);
		session.groupColumns(t, 3, lastCol);
		session.groupRows(t, 3, 3);
		if (changing.length) session.groupRows(t, changingStart, changingStart + changing.length - 1);
		if (results.length) session.groupRows(t, resultStart, resultEnd);
		session.setRowHeight(t, [3], 42);
		session.setHidden(t, 'row', [3], true);
		session.setSheetView(t, { showGridLines: false });
	});
	return created;
}
