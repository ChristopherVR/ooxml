import { compareWorkbooks } from '../xlsx/compare/compare';
import { comparisonText } from '../xlsx/compare/text';
import type { WorkbookComparison } from '../xlsx/compare/types';
import { loadXlsx } from '../xlsx/read/load';
import { saveXlsx } from '../xlsx/write/save';
import { checkEditProposal, createEditProposal, type EditProposal } from './proposal';
import { applyXlsxEdits, type XlsxEdit } from './xlsx-edits';

export type XlsxEditProposal = EditProposal<XlsxEdit>;

export interface XlsxReviewOptions {
	/** Differences to report at most (default 1000); the rest only set `truncated`. */
	maxDifferences?: number;
}

/** What a set of edits changes, as {@link compareWorkbooks} sees it, with its text report. */
export interface XlsxReview {
	comparison: WorkbookComparison;
	/** {@link comparisonText} of the comparison, for a person or an agent to read. */
	report: string;
	warnings: string[];
}

async function review(bytes: Uint8Array, edits: readonly XlsxEdit[], options: XlsxReviewOptions) {
	const before = await loadXlsx(bytes);
	const after = await loadXlsx(bytes);
	applyXlsxEdits(after, edits);
	const comparison = compareWorkbooks(before, after, {
		maxDifferences: options.maxDifferences ?? 1000,
	});
	return { after, comparison, report: comparisonText(comparison), warnings: after.warnings };
}

/**
 * Proposes `edits` without saving anything: the edits are applied to a copy of the workbook and
 * compared with it, so the result says what they would change (cells and structure, with what the
 * comparison does not look at listed in `notCompared`). Pass the proposal to
 * {@link applyXlsxProposal} once the change is approved; to accept only some edits, propose those
 * again, since edits can depend on each other.
 */
export async function proposeXlsxEdits(
	bytes: Uint8Array,
	edits: readonly XlsxEdit[],
	options: XlsxReviewOptions = {},
): Promise<XlsxReview & { proposal: XlsxEditProposal }> {
	const { comparison, report, warnings } = await review(bytes, edits, options);
	return { proposal: createEditProposal('xlsx', bytes, edits), comparison, report, warnings };
}

/**
 * Applies an approved proposal and returns the saved workbook with the comparison of what was
 * applied. It throws, changing nothing, when `bytes` are not the workbook the proposal was made
 * against or the proposal was altered after it was made.
 */
export async function applyXlsxProposal(
	bytes: Uint8Array,
	proposal: XlsxEditProposal,
	options: XlsxReviewOptions = {},
): Promise<XlsxReview & { bytes: Uint8Array }> {
	checkEditProposal(proposal, 'xlsx', bytes);
	const { after, comparison, report, warnings } = await review(bytes, proposal.edits, options);
	return { bytes: await saveXlsx(after), comparison, report, warnings };
}
