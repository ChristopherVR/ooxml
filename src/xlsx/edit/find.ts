import { rangeContains } from '../address.js';
import { forEachCell, getCell } from '../cells.js';
import type { Workbook } from '../model.js';
import { syncTableHeader, writeInput } from './cell-values.js';
import { type EditContext, displayText, sheetAt } from './context.js';
import { isSpilledCell } from './deps.js';
import type { EditScope } from './history.js';
import { formulaBarText } from './input-text.js';
import type { FindMatch, FindQuery } from './types.js';

/** A regular expression for a find query (Excel wildcards `*`, `?`, `~` unless turned off). */
export function queryPattern(query: FindQuery, global = false): RegExp | undefined {
	if (!query.text) return undefined;
	let source = '';
	const wildcards = query.wildcards ?? true;
	for (let i = 0; i < query.text.length; i++) {
		const ch = query.text[i] ?? '';
		if (wildcards && ch === '~' && i + 1 < query.text.length) {
			source += escapeRegex(query.text[++i] ?? '');
		} else if (wildcards && ch === '*') source += '[\\s\\S]*';
		else if (wildcards && ch === '?') source += '[\\s\\S]';
		else source += escapeRegex(ch);
	}
	if (query.wholeCell) source = `^(?:${source})$`;
	return new RegExp(source, `${query.matchCase ? '' : 'i'}${global ? 'g' : ''}`);
}

const escapeRegex = (text: string): string => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

function searchText(
	workbook: Workbook,
	query: FindQuery,
	sheet: number,
	row: number,
	col: number,
): string {
	const ws = workbook.sheets[sheet];
	if (!ws) return '';
	if (query.lookIn === 'comments')
		return ws.comments.find((c) => c.address.row === row && c.address.col === col)?.text ?? '';
	const cell = getCell(ws, row, col);
	return query.lookIn === 'values'
		? displayText(workbook, cell)
		: formulaBarText(workbook, cell, { quote: false });
}

/** Every cell matching the query, sheet by sheet, in row (or column) order. */
export function findAll(workbook: Workbook, query: FindQuery): FindMatch[] {
	const pattern = queryPattern(query);
	if (!pattern) return [];
	const indices = query.sheet !== undefined ? [query.sheet] : workbook.sheets.map((_s, i) => i);
	const matches: FindMatch[] = [];
	for (const s of indices) {
		const sheet = workbook.sheets[s];
		if (!sheet) continue;
		const positions: [number, number][] = [];
		if (query.lookIn === 'comments')
			for (const c of sheet.comments) positions.push([c.address.row, c.address.col]);
		else forEachCell(sheet, (_cell, row, col) => positions.push([row, col]));
		if (query.order === 'columns') positions.sort((a, b) => a[1] - b[1] || a[0] - b[0]);
		else positions.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
		for (const [row, col] of positions) {
			if (query.range && query.sheet !== undefined && !rangeContains(query.range, { row, col }))
				continue;
			const text = searchText(workbook, query, s, row, col);
			if (text && pattern.test(text)) matches.push({ sheet: s, row, col, text });
		}
	}
	return matches;
}

/** `text` with every match of the query replaced (the whole text for a whole-cell query). */
export function replaceText(text: string, query: FindQuery, replacement: string): string {
	const pattern = queryPattern(query, true);
	if (!pattern) return text;
	return text.replace(pattern, (match) => (match === '' ? '' : replacement));
}

/** Dynamic-array results are recreated by their anchor, so replacing never edits them. */
const isSpilledAt = (
	workbook: Workbook,
	at: { sheet: number; row: number; col: number },
): boolean => {
	const sheet = workbook.sheets[at.sheet];
	return !!sheet && isSpilledCell(getCell(sheet, at.row, at.col));
};

/** The query used for replacing: comments are edited as such, everything else as typed input. */
const replaceQuery = (query: FindQuery): FindQuery =>
	query.lookIn === 'comments' ? query : { ...query, lookIn: 'formulas' };

/** What replacing changes on a sheet: the matched cells (and table headers) or the comments. */
function replaceScopes(query: FindQuery, sheet: number, matches: FindMatch[]): EditScope[] {
	if (query.lookIn === 'comments') return [{ kind: 'parts', sheet, parts: ['comments'] }];
	const ranges = matches.map((m) => ({
		start: { row: m.row, col: m.col },
		end: { row: m.row, col: m.col },
	}));
	return [
		{ kind: 'cells', sheet, ranges },
		{ kind: 'parts', sheet, parts: ['tables'] },
	];
}

function applyReplace(
	ctx: EditContext,
	query: FindQuery,
	replacement: string,
	matches: FindMatch[],
): number {
	let changed = 0;
	for (const match of matches) {
		const sheet = sheetAt(ctx.workbook, match.sheet);
		const next = replaceText(match.text, query, replacement);
		if (next === match.text) continue;
		if (query.lookIn === 'comments') {
			const comment = sheet.comments.find(
				(c) => c.address.row === match.row && c.address.col === match.col,
			);
			if (comment) comment.text = next;
		} else {
			writeInput(ctx, sheet, match.row, match.col, next);
			syncTableHeader(ctx, sheet, match.row, match.col);
		}
		changed++;
	}
	return changed;
}

/**
 * Replaces the query in every matching cell and returns how many cells changed. Cells are matched
 * and edited on their formula-bar text (formulas, dates as `3/15/2023`, percents as `40%`), then
 * re-parsed like typing, as Excel does: replacing `4` never touches a date's serial.
 */
export function replaceAll(ctx: EditContext, query: FindQuery, replacement: string): number {
	const q = replaceQuery(query);
	const matches = findAll(ctx.workbook, q).filter(
		(m) => q.lookIn === 'comments' || !isSpilledAt(ctx.workbook, m),
	);
	if (!matches.length) return 0;
	const sheets = [...new Set(matches.map((m) => m.sheet))];
	const scopes = sheets.flatMap((sheet) =>
		replaceScopes(
			q,
			sheet,
			matches.filter((m) => m.sheet === sheet),
		),
	);
	return ctx.run('Replace', 'cells', scopes, () => applyReplace(ctx, q, replacement, matches), {
		...(sheets.length === 1 && sheets[0] !== undefined ? { sheet: sheets[0] } : {}),
	});
}

/** Replaces the query in one cell; false when that cell no longer matches. */
export function replaceOne(
	ctx: EditContext,
	query: FindQuery,
	replacement: string,
	at: FindMatch,
): boolean {
	const q = replaceQuery(query);
	if (q.lookIn !== 'comments' && isSpilledAt(ctx.workbook, at)) return false;
	const pattern = queryPattern(q);
	const text = searchText(ctx.workbook, q, at.sheet, at.row, at.col);
	if (!pattern || !text || !pattern.test(text)) return false;
	const match: FindMatch = { ...at, text };
	if (replaceText(text, q, replacement) === text) return false;
	ctx.run(
		'Replace',
		'cells',
		replaceScopes(q, at.sheet, [match]),
		() => applyReplace(ctx, q, replacement, [match]),
		{ sheet: at.sheet },
	);
	return true;
}
