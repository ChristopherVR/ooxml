import { z } from 'zod';
import {
	createXlsx,
	findXlsxCells,
	inspectXlsx,
	readXlsxFilledCells,
	readXlsxRange,
	readXlsxTable,
	setXlsxCells,
	type XlsxFindOptions,
} from '../xlsx';
import { createTool, editTool, id, index, readTool } from './define';
import type { AutomationTool } from './types';

const format = 'xlsx';
const readable = ['.xlsx', '.xlsm'];

/**
 * Spreadsheet tools, reads first. Every read is bounded so a large workbook cannot flood the
 * caller: start with `xlsx_inspect` (sheets and their used ranges), then read a range, the filled
 * cells up to a cap, a table, or the matches of a search.
 */
export const xlsxTools: readonly AutomationTool[] = [
	readTool({
		name: 'xlsx_inspect',
		description:
			'List the sheets (index, name, state, used range and its size, tables), the active sheet, defined names, properties and warnings. Start here: it says how big each sheet is before reading cells.',
		format,
		extensions: readable,
		input: {},
		run: (bytes) => inspectXlsx(bytes),
	}),
	readTool({
		name: 'xlsx_read_range',
		description:
			'Read the cells of an A1 range (at most 10000 cells): value, formula, displayed text, number format and, for dates, ISO 8601.',
		format,
		extensions: readable,
		input: { sheetIndex: index, range: id },
		run: (bytes, p) => readXlsxRange(bytes, p.sheetIndex, p.range),
	}),
	readTool({
		name: 'xlsx_read_filled_cells',
		description:
			'Read the filled cells (values and formulas) of one sheet or every sheet in row order, up to maxCells in all (default 2000). Each sheet reports its total and how many were omitted.',
		format,
		extensions: readable,
		input: {
			sheetIndex: index.optional(),
			maxCells: z.number().int().min(1).max(10_000).default(2000),
		},
		run: (bytes, p) =>
			readXlsxFilledCells(bytes, {
				maxCells: p.maxCells,
				...(p.sheetIndex !== undefined && { sheetIndex: p.sheetIndex }),
			}),
	}),
	readTool({
		name: 'xlsx_find',
		description:
			"Find cells the way Excel's Find does (wildcards * ? ~ on by default), in one sheet or range or the whole workbook. Returns at most maxResults matches (default 100) and the total.",
		format,
		extensions: readable,
		input: {
			text: id.max(32_767),
			sheetIndex: index.optional(),
			range: id.optional(),
			lookIn: z.enum(['formulas', 'values']).optional(),
			matchCase: z.boolean().optional(),
			wholeCell: z.boolean().optional(),
			wildcards: z.boolean().optional(),
			maxResults: z.number().int().min(1).max(1000).optional(),
		},
		run: (bytes, { text, ...options }) => {
			const defined: XlsxFindOptions = {};
			for (const [key, value] of Object.entries(options))
				if (value !== undefined) Object.assign(defined, { [key]: value });
			return findXlsxCells(bytes, text, defined);
		},
	}),
	readTool({
		name: 'xlsx_read_table',
		description:
			'Read a table by name (case-insensitive): its sheet, range, header and totals rows, column names and cells (at most 10000 cells).',
		format,
		extensions: readable,
		input: { name: id },
		run: (bytes, p) => readXlsxTable(bytes, p.name),
	}),
	createTool({
		name: 'xlsx_create',
		description: 'Create a workbook with the named sheets.',
		format,
		extensions: ['.xlsx'],
		input: { sheets: z.array(id).max(100).default([]) },
		run: (p) => createXlsx(p.sheets),
	}),
	editTool({
		name: 'xlsx_set_cells',
		description:
			'Set cell inputs through core editing and formula recalculation (at most 10000 cells). Inputs beginning with = are formulas.',
		format,
		extensions: readable,
		input: {
			sheetIndex: index,
			cells: z
				.array(z.strictObject({ address: id, input: z.string().max(32_767) }))
				.min(1)
				.max(10_000),
		},
		run: (bytes, p) => setXlsxCells(bytes, p.sheetIndex, p.cells),
	}),
];
