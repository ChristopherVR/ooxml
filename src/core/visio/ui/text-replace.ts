import type { VisioDocument } from '../model';
import { replaceEntries, snapshotReplaceScope, type ReplaceEntry } from './text-replace-snapshot';
import {
	VISIO_TEXT_REPLACE_LIMITS as limits,
	replaceIdentity,
	replaceString,
	textReplaceError,
	type VisioTextOccurrence,
	type VisioTextReplacePlan,
	type VisioTextReplaceRequest,
	type VisioTextReplaceScope,
	type VisioTextReplaceEdit,
} from './text-replace-types';
export type {
	VisioTextOccurrence,
	VisioTextReplacePlan,
	VisioTextReplaceRequest,
	VisioTextReplaceScope,
	VisioTextReplaceEdit,
	VisioTextReplacePlanEdit,
} from './text-replace-types';
export { VISIO_TEXT_REPLACE_LIMITS } from './text-replace-types';
export { visioTextOccurrenceSelection } from './text-replace-selection';

function occurrences(
	entries: readonly ReplaceEntry[],
	query: string,
): readonly VisioTextOccurrence[] {
	const result: VisioTextOccurrence[] = [];
	for (const entry of entries) {
		let start = entry.text.indexOf(query);
		while (start >= 0) {
			if (result.length >= limits.matches)
				textReplaceError('Replace occurrence count exceeds limits.');
			result.push(
				Object.freeze({
					pageId: entry.pageId,
					shapeId: entry.shapeId,
					start,
					end: start + query.length,
				}),
			);
			start = entry.text.indexOf(query, start + query.length);
		}
	}
	return Object.freeze(result);
}
/** Full literal occurrences, independent of the bounded preview-oriented Find index. */
export function visioTextReplaceOccurrences(
	model: VisioDocument,
	input: VisioTextReplaceScope,
): readonly VisioTextOccurrence[] {
	const scope = snapshotReplaceScope(input);
	return occurrences(replaceEntries(model, scope), scope.query);
}

/** Next occurrence in bounded source order, wrapping after the last occurrence. */
export function visioTextReplaceFindNext(
	model: VisioDocument,
	input: VisioTextReplaceScope,
	current?: VisioTextOccurrence,
): VisioTextOccurrence | undefined {
	const saved = snapshotOccurrence(current);
	const found = visioTextReplaceOccurrences(model, input);
	if (saved === undefined) return found[0];
	const index = found.findIndex((match) => sameOccurrence(match, saved));
	if (index < 0) return textReplaceError('Find occurrence is stale or outside the scope.');
	return found[(index + 1) % found.length];
}

const records = new WeakMap<
	VisioTextReplacePlan,
	{ scope: VisioTextReplaceScope; entries: ReplaceEntry[] }
>();
const sameOccurrence = (a: VisioTextOccurrence, b: VisioTextOccurrence): boolean =>
	a.pageId === b.pageId && a.shapeId === b.shapeId && a.start === b.start && a.end === b.end;
const sameTarget = (entry: ReplaceEntry, match: VisioTextOccurrence): boolean =>
	entry.pageId === match.pageId && entry.shapeId === match.shapeId;

function snapshotOccurrence(
	original: VisioTextOccurrence | undefined,
): VisioTextOccurrence | undefined {
	if (original === undefined) return undefined;
	const current = {
		pageId: replaceIdentity(original.pageId),
		shapeId: replaceIdentity(original.shapeId),
		start: original.start,
		end: original.end,
	};
	if (
		!Number.isSafeInteger(current.start) ||
		!Number.isSafeInteger(current.end) ||
		current.start < 0 ||
		current.end <= current.start ||
		current.end > limits.text
	)
		textReplaceError('Invalid replacement occurrence range.');
	return current;
}

/** Selected groups include their subtree, deduplicated in ordered-selection/source order.
 * Source editing remains authoritative for rich text, fields, inheritance and protection.
 * Callers must also bind application source revisions when applying these scene-derived plans.
 */
export function visioTextReplacePlan(
	model: VisioDocument,
	input: VisioTextReplaceRequest,
): VisioTextReplacePlan {
	const scope = snapshotReplaceScope(input),
		replacement = replaceString(input.replacement, limits.replacement, 'replacement'),
		mode = input.mode,
		current = snapshotOccurrence(input.current);
	if (mode !== 'current' && mode !== 'all') textReplaceError('Invalid replacement mode.');
	if (
		/[\u0000-\u0008\u000b\u000c\u000d\u000e-\u001f\ufffe\uffff]|[\ud800-\udbff](?![\udc00-\udfff])|(?<![\ud800-\udbff])[\udc00-\udfff]/u.test(
			replacement,
		)
	)
		textReplaceError('Replacement contains invalid XML text characters.');
	const entries = replaceEntries(model, scope),
		found = occurrences(entries, scope.query);
	const active =
		current === undefined ? found[0] : found.find((match) => sameOccurrence(match, current));
	if (current !== undefined && !active)
		textReplaceError('Replace occurrence is stale or outside the scope.');
	const selected = mode === 'all' ? found : active ? [active] : [];
	const targets = new Map<string, VisioTextOccurrence[]>();
	for (const match of selected) {
		const key = JSON.stringify([match.pageId, match.shapeId]);
		const list = targets.get(key) ?? [];
		list.push(match);
		targets.set(key, list);
	}
	if (targets.size > limits.commands) textReplaceError('Replace command count exceeds limits.');
	const edits: VisioTextReplaceEdit[] = [],
		updated: ReplaceEntry[] = [];
	let outputCharacters = 0,
		commandCharacters = 0;
	for (const entry of entries) {
		const matches = targets.get(JSON.stringify([entry.pageId, entry.shapeId]));
		let text = entry.text;
		if (matches) {
			const length = text.length + matches.length * (replacement.length - scope.query.length);
			outputCharacters += length;
			if (length > limits.output || outputCharacters > limits.output)
				textReplaceError('Replace output exceeds aggregate command text limits.');
			const parts: string[] = [];
			let from = 0;
			for (const match of matches) {
				parts.push(text.slice(from, match.start), replacement);
				from = match.end;
			}
			parts.push(text.slice(from));
			text = parts.join('');
			// Retain matched no-op targets so unsupported/protected text is never silently skipped.
			const target = { pageId: entry.pageId, shapeId: entry.shapeId };
			if (replacement.includes('\n') || scope.query.includes('\n')) {
				commandCharacters += text.length;
				edits.push(Object.freeze({ type: 'replace-plain-text', ...target, text }));
			} else {
				commandCharacters += entry.text.length + matches.length * replacement.length;
				edits.push(
					Object.freeze({
						type: 'replace-text-ranges',
						...target,
						expectedText: entry.text,
						ranges: Object.freeze(
							matches.map((match) =>
								Object.freeze({ start: match.start, end: match.end, text: replacement }),
							),
						),
					}),
				);
			}
			if (commandCharacters > limits.output)
				textReplaceError('Replace input exceeds aggregate command text limits.');
		}
		updated.push({ ...entry, text });
	}
	const remaining = mode === 'current' ? occurrences(updated, scope.query) : [];
	const activeEntry = active ? entries.findIndex((entry) => sameTarget(entry, active)) : -1;
	const entryIndices = new Map(
		entries.map((entry, index) => [JSON.stringify([entry.pageId, entry.shapeId]), index]),
	);
	const nextOccurrence = active
		? (remaining.find((match) => {
				const index = entryIndices.get(JSON.stringify([match.pageId, match.shapeId]))!;
				return (
					index > activeEntry ||
					(index === activeEntry && match.start >= active.start + replacement.length)
				);
			}) ?? remaining[0])
		: remaining[0];
	const plan = Object.freeze({
		occurrences: found,
		replacementCount: selected.length,
		edits: Object.freeze(edits),
		nextOccurrence,
	});
	records.set(plan, { scope, entries });
	return plan;
}

/** Reject forged and stale scene plans before returning independently owned atomic commands. */
export function visioTextReplaceCommands(
	model: VisioDocument,
	plan: VisioTextReplacePlan,
): VisioTextReplaceEdit[] {
	const record = records.get(plan);
	if (!record) return textReplaceError('Replace plan was not produced by this planner.');
	const current = replaceEntries(model, record.scope);
	if (
		current.length !== record.entries.length ||
		current.some((entry, index) => {
			const saved = record.entries[index]!;
			return (
				entry.pageId !== saved.pageId ||
				entry.shapeId !== saved.shapeId ||
				entry.text !== saved.text
			);
		})
	)
		textReplaceError('Replace plan is stale. Search the current drawing again.');
	return plan.edits.map((edit) =>
		edit.type === 'replace-text-ranges'
			? { ...edit, ranges: edit.ranges.map((range) => ({ ...range })) }
			: { ...edit },
	);
}
