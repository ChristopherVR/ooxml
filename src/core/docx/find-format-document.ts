// Document-wide formatted Replace across body, headers, footers, footnotes and endnotes.
import { findFormattedMatches, modelFormatContext } from './find-format';
import {
	cloneBlocks,
	replaceFormattedMatches,
	type ReplaceFormatOptions,
	type ReplaceFormatResult,
} from './find-format-replace';
import type { Block, DocumentModel } from './model';

export interface DocumentReplaceFormatResult {
	readonly model: DocumentModel;
	readonly replacedCount: number;
}

/** Replaces matched text and/or formatting across the body, headers, footers and notes of `model`. */
export function replaceDocumentFormattedMatches(
	model: DocumentModel,
	query: string,
	replacement: string,
	options: ReplaceFormatOptions = {},
): DocumentReplaceFormatResult {
	const context = modelFormatContext(model, options.context);
	const baseOptions: ReplaceFormatOptions = { ...options, context };
	let replacedCount = 0;
	let remainingIndex = options.matchIndex;

	const replaceInBlocks = (blocks: readonly Block[]): ReplaceFormatResult => {
		if (remainingIndex === undefined) {
			const res = replaceFormattedMatches(blocks, query, replacement, baseOptions);
			replacedCount += res.replacedCount;
			return res;
		}
		if (remainingIndex < 0) return { blocks: cloneBlocks(blocks), replacedCount: 0 };
		const found = findFormattedMatches(blocks, query, baseOptions);
		if (remainingIndex >= found.length) {
			remainingIndex -= found.length;
			return { blocks: cloneBlocks(blocks), replacedCount: 0 };
		}
		const targetIndex = remainingIndex;
		remainingIndex = -1;
		const res = replaceFormattedMatches(blocks, query, replacement, {
			...baseOptions,
			matchIndex: targetIndex,
		});
		replacedCount += res.replacedCount;
		return res;
	};

	const bodyResult = replaceInBlocks(model.blocks);
	const seenSlots = new Map<object, { blocks: Block[] }>();

	const replaceSlot = <T extends { blocks: Block[] }>(slot: T): T => {
		const existing = seenSlots.get(slot);
		if (existing) return existing as T;
		const res = replaceInBlocks(slot.blocks);
		const nextSlot = res.replacedCount > 0 ? { ...slot, blocks: res.blocks } : slot;
		seenSlots.set(slot, nextSlot);
		return nextSlot;
	};

	const replaceSlots = (slots: {
		default?: { blocks: Block[] };
		first?: { blocks: Block[] };
		even?: { blocks: Block[] };
	}) => ({
		...(slots.default ? { default: replaceSlot(slots.default) } : {}),
		...(slots.first ? { first: replaceSlot(slots.first) } : {}),
		...(slots.even ? { even: replaceSlot(slots.even) } : {}),
	});

	const sections = model.sections?.map((section) => ({
		...section,
		...(section.headers ? { headers: replaceSlots(section.headers) } : {}),
		...(section.footers ? { footers: replaceSlots(section.footers) } : {}),
	}));

	const replaceNotes = (notes: DocumentModel['footnotes']) =>
		notes?.map((note) => {
			const res = replaceInBlocks(note.blocks);
			return res.replacedCount > 0 ? { ...note, blocks: res.blocks } : note;
		});

	const footnotes = replaceNotes(model.footnotes);
	const endnotes = replaceNotes(model.endnotes);

	const nextModel: DocumentModel = {
		...model,
		blocks: bodyResult.blocks,
		...(sections !== undefined ? { sections } : {}),
		...(footnotes !== undefined ? { footnotes } : {}),
		...(endnotes !== undefined ? { endnotes } : {}),
	};
	return { model: nextModel, replacedCount };
}
