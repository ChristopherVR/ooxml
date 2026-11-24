import type { CellRange } from '../address';
import { getCell } from '../cells';
import type { Cell, Workbook, Worksheet } from '../model';
import { displayText } from './context';

type CompareOp = '=' | '<>' | '>' | '>=' | '<' | '<=';

interface ParsedCriterion {
	op: CompareOp;
	operand: string;
	numOperand: number;
	wildcard: RegExp | undefined;
}

type CellPredicate = (cell: Cell | undefined) => boolean;

function compileWildcard(pattern: string): RegExp | undefined {
	let out = '';
	let hasWild = false;
	for (let i = 0; i < pattern.length; i++) {
		const ch = pattern[i]!;
		if (ch === '~' && i + 1 < pattern.length) {
			hasWild = true;
			out += pattern[++i]!.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
		} else if (ch === '*') {
			hasWild = true;
			out += '[\\s\\S]*';
		} else if (ch === '?') {
			hasWild = true;
			out += '[\\s\\S]';
		} else {
			out += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
		}
	}
	return hasWild ? new RegExp(`^${out}$`, 'i') : undefined;
}

function parseCriterion(raw: string): ParsedCriterion {
	let op: CompareOp = '=';
	let operand = raw;
	if (raw.startsWith('<>')) {
		op = '<>';
		operand = raw.slice(2);
	} else if (raw.startsWith('>=')) {
		op = '>=';
		operand = raw.slice(2);
	} else if (raw.startsWith('<=')) {
		op = '<=';
		operand = raw.slice(2);
	} else if (raw.startsWith('>')) {
		op = '>';
		operand = raw.slice(1);
	} else if (raw.startsWith('<')) {
		op = '<';
		operand = raw.slice(1);
	} else if (raw.startsWith('=')) {
		op = '=';
		operand = raw.slice(1);
	}
	const trimmed = operand.trim();
	const numOperand = trimmed === '' ? Number.NaN : Number(trimmed);
	const wildcard = op === '=' || op === '<>' ? compileWildcard(operand) : undefined;
	return { op, operand, numOperand, wildcard };
}

function isBlankCell(cell: Cell | undefined): boolean {
	return cell === undefined || cell.value === null || cell.value === '';
}

function compareCell(workbook: Workbook, cell: Cell | undefined, parsed: ParsedCriterion): number {
	const value = cell?.value ?? null;
	if (typeof value === 'number' && Number.isFinite(parsed.numOperand))
		return value - parsed.numOperand;
	if (typeof value === 'boolean') {
		const lower = parsed.operand.trim().toLowerCase();
		if (lower === 'true' || lower === 'false') {
			const b = lower === 'true';
			return value === b ? 0 : value ? 1 : -1;
		}
	}
	const a = displayText(workbook, cell).toLowerCase();
	const b = parsed.operand.toLowerCase();
	return a < b ? -1 : a > b ? 1 : 0;
}

function compileCellPredicate(workbook: Workbook, raw: string): CellPredicate {
	const parsed = parseCriterion(raw);
	return (cell: Cell | undefined): boolean => {
		const blank = isBlankCell(cell);
		if (parsed.operand === '') {
			if (parsed.op === '=') return blank;
			if (parsed.op === '<>') return !blank;
			return false;
		}
		if (blank) return parsed.op === '<>';
		if (parsed.wildcard) {
			const hit = parsed.wildcard.test(displayText(workbook, cell));
			return parsed.op === '=' ? hit : !hit;
		}
		const cmp = compareCell(workbook, cell, parsed);
		switch (parsed.op) {
			case '=':
				return cmp === 0;
			case '<>':
				return cmp !== 0;
			case '>':
				return cmp > 0;
			case '>=':
				return cmp >= 0;
			case '<':
				return cmp < 0;
			case '<=':
				return cmp <= 0;
		}
	};
}

/**
 * Compiles an Advanced Filter criteria range (header row + 1+ criteria rows) into a row predicate.
 * Columns within the same criteria row are ANDed; multiple criteria rows are ORed.
 */
export function compileAdvancedCriteria(
	workbook: Workbook,
	listSheet: Worksheet,
	listRange: CellRange,
	criteriaSheet: Worksheet,
	criteriaRange: CellRange,
): (row: number) => boolean {
	if (criteriaRange.end.row <= criteriaRange.start.row) return () => true;

	const headerToCol = new Map<string, number>();
	for (let col = listRange.start.col; col <= listRange.end.col; col++) {
		const name = displayText(workbook, getCell(listSheet, listRange.start.row, col))
			.trim()
			.toLowerCase();
		if (name && !headerToCol.has(name)) headerToCol.set(name, col);
	}

	const criteriaCols: (number | undefined)[] = [];
	for (let col = criteriaRange.start.col; col <= criteriaRange.end.col; col++) {
		const name = displayText(workbook, getCell(criteriaSheet, criteriaRange.start.row, col))
			.trim()
			.toLowerCase();
		criteriaCols.push(name ? headerToCol.get(name) : undefined);
	}

	const rowTests: ((row: number) => boolean)[] = [];
	for (let r = criteriaRange.start.row + 1; r <= criteriaRange.end.row; r++) {
		const andParts: ((row: number) => boolean)[] = [];
		for (let i = 0; i < criteriaCols.length; i++) {
			const critCol = criteriaRange.start.col + i;
			const critCell = getCell(criteriaSheet, r, critCol);
			const raw = displayText(workbook, critCell);
			if (raw === '') continue;
			const listCol = criteriaCols[i];
			if (listCol === undefined) {
				andParts.push(() => false);
				continue;
			}
			const pred = compileCellPredicate(workbook, raw);
			andParts.push((dataRow) => pred(getCell(listSheet, dataRow, listCol)));
		}
		rowTests.push(andParts.length ? (dataRow) => andParts.every((fn) => fn(dataRow)) : () => true);
	}

	if (!rowTests.length) return () => true;
	return (dataRow: number) => rowTests.some((fn) => fn(dataRow));
}
