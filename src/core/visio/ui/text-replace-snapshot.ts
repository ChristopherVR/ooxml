import type { VisioDocument, VisioShape } from '../model';
import {
	VISIO_TEXT_REPLACE_LIMITS as limits,
	replaceIdentity,
	replaceString,
	textReplaceError,
	type VisioTextReplaceScope,
} from './text-replace-types';

export interface ReplaceEntry {
	pageId: string;
	shapeId: string;
	name: string;
	pageIndex: number;
	text: string;
	end: number;
}
export function snapshotReplaceScope(input: VisioTextReplaceScope): VisioTextReplaceScope {
	const query = replaceString(input.query, limits.query, 'query');
	const scope = input.scope,
		pageId = replaceIdentity(input.pageId),
		matchCase = input.matchCase;
	if (!query || matchCase !== true || !['selection', 'current-page', 'all-pages'].includes(scope))
		textReplaceError(
			'Replace requires a nonempty literal query, Match case, and an explicit scope.',
		);
	let selection: { id: string; pageId?: string }[] | undefined;
	const original = input.selection;
	if (original !== undefined) {
		const length = original.length;
		if (
			!Array.isArray(original) ||
			!Number.isSafeInteger(length) ||
			length > limits.shapes ||
			length < 0
		)
			textReplaceError('Replace selection exceeds supported limits.');
		selection = [];
		for (let index = 0; index < length; index++) {
			const item = original[index]!,
				id = replaceIdentity(item.id),
				page = item.pageId;
			selection.push(
				Object.freeze({ id, ...(page === undefined ? {} : { pageId: replaceIdentity(page) }) }),
			);
		}
		Object.freeze(selection);
	}
	return Object.freeze({ query, scope, pageId, matchCase, ...(selection ? { selection } : {}) });
}

/** Own only text and identities; no geometry, previews, host iterators or mutable shape records. */
export function replaceEntries(
	model: VisioDocument,
	scope: VisioTextReplaceScope,
	requirePage = true,
): ReplaceEntry[] {
	const pages = model.pages;
	const count = pages.length;
	if (!Array.isArray(pages) || !Number.isSafeInteger(count) || count > limits.pages || count < 0)
		textReplaceError('Replace page list exceeds supported limits.');
	const all: ReplaceEntry[] = [],
		identities = new Map<string, number>(),
		pageIds = new Set<string>(),
		seen = new Set<VisioShape>();
	let characters = 0,
		names = 0;
	for (let pageIndex = 0; pageIndex < count; pageIndex++) {
		const page = pages[pageIndex]!,
			pageId = replaceIdentity(page.id),
			ids = new Set<string>();
		if (pageIds.has(pageId)) textReplaceError('Replace page identities are ambiguous.');
		pageIds.add(pageId);
		const walk = (source: readonly VisioShape[], depth: number): void => {
			const length = source.length;
			if (
				!Array.isArray(source) ||
				!Number.isSafeInteger(length) ||
				length < 0 ||
				(length > 0 && depth > limits.depth) ||
				length > limits.shapes
			)
				textReplaceError('Replace shape list exceeds supported limits.');
			for (let index = 0; index < length; index++) {
				const shape = source[index]!,
					shapeId = replaceIdentity(shape.id);
				if (seen.has(shape) || ids.has(shapeId) || seen.size >= limits.shapes)
					textReplaceError('Replace shape identities, depth or cycles are invalid.');
				seen.add(shape);
				ids.add(shapeId);
				const text = replaceString(shape.text.plainText, limits.text, 'source text');
				const name = replaceString(shape.name, 4096, 'shape name');
				names += name.length;
				if (names > limits.names) textReplaceError('Replace shape names exceed aggregate limits.');
				characters += text.length;
				if (characters > limits.text) textReplaceError('Replace text exceeds aggregate limits.');
				const entry = { pageId, shapeId, name, pageIndex, text, end: 0 };
				identities.set(JSON.stringify([pageId, shapeId]), all.length);
				all.push(entry);
				walk(shape.children, depth + 1);
				entry.end = all.length;
			}
		};
		walk(page.shapes, 0);
	}
	if (requirePage && !pageIds.has(scope.pageId))
		textReplaceError('Replace current page does not exist.');
	if (scope.scope === 'all-pages') return all;
	if (scope.scope === 'current-page') return all.filter((entry) => entry.pageId === scope.pageId);
	const result: ReplaceEntry[] = [],
		added = new Set<ReplaceEntry>();
	for (const selected of scope.selection ?? []) {
		const pageId = selected.pageId ?? scope.pageId;
		const start = identities.get(JSON.stringify([pageId, selected.id]));
		if (start === undefined) return textReplaceError('Replace selected shape does not exist.');
		for (let index = start; index < all[start]!.end; index++) {
			const entry = all[index]!;
			if (added.has(entry)) {
				index = entry.end - 1;
				continue;
			}
			result.push(entry);
			added.add(entry);
		}
	}
	return result;
}
