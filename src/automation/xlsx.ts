import { createWorkbook, validateSheetName } from '../xlsx/workbook.js';
import { loadXlsx } from '../xlsx/read/load.js';
import { saveXlsx } from '../xlsx/write/save.js';
import { createEditSession } from '../xlsx/edit/session.js';
import { parseAddress, parseRange, MAX_COL, MAX_ROW, formatAddress } from '../xlsx/address.js';
import { getCell } from '../xlsx/cells.js';

export async function inspectXlsx(bytes: Uint8Array) {
	const workbook = await loadXlsx(bytes);
	return {
		sheets: workbook.sheets.map((sheet, index) => ({ index, name: sheet.name })),
		properties: workbook.properties,
		warnings: workbook.warnings,
	};
}

export async function createXlsx(names: readonly string[]): Promise<Uint8Array> {
	const workbook = createWorkbook();
	const session = createEditSession(workbook);
	if (names.length) {
		const first = names[0]!;
		const problem = validateSheetName(workbook, first, 0);
		if (problem) throw new Error(problem);
		session.renameSheet(0, first);
		for (const name of names.slice(1)) session.addSheet(name);
	}
	return saveXlsx(workbook);
}

export async function readXlsxRange(bytes: Uint8Array, sheetIndex: number, reference: string) {
	const workbook = await loadXlsx(bytes);
	const sheet =
		Number.isSafeInteger(sheetIndex) && sheetIndex >= 0 ? workbook.sheets[sheetIndex] : undefined;
	const range = parseRange(reference);
	if (!sheet || !range || range.end.col > MAX_COL || range.end.row > MAX_ROW)
		throw new Error('Invalid sheet or range');
	const count = (range.end.row - range.start.row + 1) * (range.end.col - range.start.col + 1);
	if (count > 10_000) throw new Error('Read at most 10000 cells per request');
	const cells = [];
	for (let row = range.start.row; row <= range.end.row; row++)
		for (let col = range.start.col; col <= range.end.col; col++) {
			const cell = getCell(sheet, row, col);
			cells.push({
				address: formatAddress({ row, col }),
				value: cell?.value ?? null,
				formula: cell?.formula,
			});
		}
	return { cells, warnings: workbook.warnings };
}

export async function setXlsxCells(
	bytes: Uint8Array,
	sheetIndex: number,
	cells: readonly { address: string; input: string }[],
) {
	const workbook = await loadXlsx(bytes);
	if (!Number.isSafeInteger(sheetIndex) || sheetIndex < 0 || !workbook.sheets[sheetIndex])
		throw new Error('Sheet does not exist');
	if (cells.length > 10_000) throw new Error('Edit at most 10000 cells per request');
	const positions = cells.map((cell) => {
		const address = parseAddress(cell.address);
		if (!address) throw new Error(`Invalid cell address: ${cell.address}`);
		return { ...address, input: cell.input };
	});
	const session = createEditSession(workbook);
	session.batch('Automation cell edits', () => {
		for (const cell of positions) session.setCellInput(sheetIndex, cell.row, cell.col, cell.input);
	});
	return { bytes: await saveXlsx(workbook), warnings: workbook.warnings };
}
