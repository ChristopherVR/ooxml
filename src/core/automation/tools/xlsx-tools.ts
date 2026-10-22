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
import { XLSX_EDIT_LIMITS, type XlsxEdit } from '../xlsx-edits';
import { applyXlsxProposal, proposeXlsxEdits, type XlsxReview } from '../xlsx-proposal';
import { createTool, editTool, id, index, readTool } from './define';
import type { AutomationTool } from './types';

const format = 'xlsx';
const readable = ['.xlsx', '.xlsm'];

const sheet = index;
const range = id;
const count = z.number().int().min(1);
const cellInputs = z
	.array(z.strictObject({ address: id, input: z.string().max(32_767) }))
	.min(1)
	.max(10_000);
/** The edits of a proposal (`XlsxEdit`); `applyXlsxEdits` checks them again in core. */
const xlsxEdit = z.discriminatedUnion('op', [
	z.strictObject({ op: z.literal('setCells'), sheet, cells: cellInputs }),
	z.strictObject({
		op: z.literal('clear'),
		sheet,
		range,
		what: z.enum(['all', 'contents', 'formats', 'comments', 'hyperlinks']).optional(),
	}),
	z.strictObject({
		op: z.literal('format'),
		sheet,
		range,
		style: z
			.strictObject({
				numFmt: z.string().max(255).optional(),
				font: z.record(z.string(), z.unknown()).optional(),
				fill: z.record(z.string(), z.unknown()).optional(),
				border: z.record(z.string(), z.unknown()).optional(),
				alignment: z.record(z.string(), z.unknown()).optional(),
				protection: z.record(z.string(), z.unknown()).optional(),
			})
			.describe('A style patch: numFmt, font, fill, border, alignment, protection'),
	}),
	z.strictObject({
		op: z.literal('merge'),
		sheet,
		range,
		mode: z.enum(['merge', 'center', 'across']).optional(),
	}),
	z.strictObject({ op: z.literal('unmerge'), sheet, range }),
	z.strictObject({
		op: z.enum(['insertRows', 'deleteRows']),
		sheet,
		row: z.number().int().min(1).describe('Row number, from 1'),
		count,
	}),
	z.strictObject({
		op: z.enum(['insertColumns', 'deleteColumns']),
		sheet,
		column: z
			.string()
			.regex(/^[A-Za-z]{1,3}$/)
			.describe('Column letters'),
		count,
	}),
	z.strictObject({ op: z.literal('addSheet'), name: id.optional(), at: index.optional() }),
	z.strictObject({ op: z.literal('renameSheet'), sheet, name: id }),
	z.strictObject({ op: z.literal('deleteSheet'), sheet }),
	z.strictObject({
		op: z.literal('setName'),
		name: id,
		formula: z.string().min(1).max(8192),
		sheet: sheet.optional(),
	}),
	z.strictObject({ op: z.literal('deleteName'), name: id, sheet: sheet.optional() }),
]);
const edits = z.array(xlsxEdit).min(1).max(XLSX_EDIT_LIMITS.edits);
const maxDifferences = z.number().int().min(1).max(1000).optional();

/** The text report and counts of a review, without the per-cell styles of the full comparison. */
function reviewResult(review: XlsxReview) {
	const { comparison, report, warnings } = review;
	return {
		report,
		differences: comparison.count,
		truncated: comparison.truncated,
		notCompared: comparison.notCompared,
		warnings,
	};
}

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
	readTool({
		name: 'xlsx_propose_edits',
		description:
			'Propose edits without saving: returns what they would change (a text diff of cells, sheets, names, merges, tables, validation and conditional formats) and a proposal to pass unchanged to xlsx_apply_proposal once the user approves. Edits run in order; sheets are indexes from 0, rows numbered from 1, columns letters.',
		format,
		extensions: readable,
		input: { edits, maxDifferences },
		run: async (bytes, p) => {
			const review = await proposeXlsxEdits(
				bytes,
				p.edits as XlsxEdit[],
				p.maxDifferences !== undefined ? { maxDifferences: p.maxDifferences } : {},
			);
			return { proposal: review.proposal, ...reviewResult(review) };
		},
	}),
	editTool({
		name: 'xlsx_apply_proposal',
		description:
			'Apply a proposal from xlsx_propose_edits after the user approved its diff. Refused, changing nothing, if the workbook changed since the proposal was made or the proposal was altered.',
		format,
		extensions: readable,
		input: {
			proposal: z.strictObject({
				format: z.literal('xlsx'),
				base: z.string().regex(/^[0-9a-f]{64}$/),
				edits,
				id: z.string().regex(/^[0-9a-f]{64}$/),
			}),
		},
		run: async (bytes, p) => {
			const applied = await applyXlsxProposal(bytes, {
				...p.proposal,
				edits: p.proposal.edits as XlsxEdit[],
			});
			return { bytes: applied.bytes, ...reviewResult(applied) };
		},
	}),
];
