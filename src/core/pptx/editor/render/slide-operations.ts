/**
 * slide-operations: Pure factory helpers for blank slides + slide ids.
 *
 * Framework-agnostic (no framework imports). The React `useSlideManagement`
 * hook and the Vue `useSlideOperations` composable both call these so the
 * blank-slide shape and id scheme live in one place.
 *
 * `makeSlideId` uses the `slide-${Date.now()}-${random}` pattern inside its
 * function body (called at runtime, never at module-eval time) and accepts an
 * optional `idGenerator` override, matching the neighbouring shared editor
 * modules' id convention.
 */

import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { cloneElement, cloneSlide, duplicateSlide } from '../../index';

export interface SlidePasteRequest {
	/** Independent slide copies with their new identities already assigned. */
	readonly copies: PptxSlide[];
	/** Original clipboard identities, for the host's separate template-view cache. */
	readonly sourceIds: string[];
	/** Zero-based insertion position, immediately after the destination selection. */
	readonly insertAt: number;
	/** Template-view content captured with the clipboard, never read from a later live cache. */
	readonly templateElements: PptxElement[][];
}

export function captureSlideClipboard(
	slides: readonly PptxSlide[],
	ids: readonly string[],
): PptxSlide[] {
	return ids.flatMap((id) => {
		const slide = slides.find((candidate) => candidate.id === id);
		return slide ? [cloneSlide(slide)] : [];
	});
}

export function createSlidePaste(
	clipboard: readonly PptxSlide[],
	insertAt: number,
	templates: Record<string, PptxElement[]> = {},
): SlidePasteRequest {
	return {
		copies: clipboard.map((slide, index) => ({
			...duplicateSlide(slide, insertAt + index + 1),
			id: makeSlideId(),
		})),
		sourceIds: clipboard.map((slide) => slide.id),
		insertAt,
		templateElements: clipboard.map((slide) => (templates[slide.id] ?? []).map(cloneElement)),
	};
}

/** Apply one paste atomically; copied content and source package hints remain independent. */
export function insertSlidePaste(
	slides: readonly PptxSlide[],
	request: SlidePasteRequest,
): PptxSlide[] {
	const index = Math.max(0, Math.min(slides.length, request.insertAt));
	const next = [
		...slides.slice(0, index),
		...request.copies.map(cloneSlide),
		...slides.slice(index),
	];
	return next.map((slide, position) =>
		slide.slideNumber === position + 1 ? slide : { ...slide, slideNumber: position + 1 },
	);
}

export function applySlidePasteTemplates(
	mapping: Record<string, PptxElement[]>,
	request: SlidePasteRequest,
): Record<string, PptxElement[]> {
	const next = { ...mapping };
	request.copies.forEach((copy, index) => {
		next[copy.id] = (request.templateElements[index] ?? []).map(cloneElement);
	});
	return next;
}

export function captureSlideClipboardTemplates(
	mapping: Record<string, PptxElement[]>,
	ids: readonly string[],
): Record<string, PptxElement[]> {
	return Object.fromEntries(ids.map((id) => [id, (mapping[id] ?? []).map(cloneElement)]));
}

/**
 * Generate a collision-resistant slide id of the form
 * `slide-<timestamp>-<base36 suffix>`.
 *
 * @param idGenerator - Optional override; when supplied its result is returned
 *   verbatim (callers wanting `crypto.randomUUID` ids pass it here).
 */
export function makeSlideId(idGenerator?: () => string): string {
	if (idGenerator) {
		return idGenerator();
	}
	return `slide-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`;
}

/**
 * Build a fresh, empty {@link PptxSlide}: a new id, empty `rId`, the given
 * `slideNumber`, and an empty `elements` array. This is the minimal shape both
 * the React `handleAddSlide` and the Vue `createBlankSlide` produced.
 *
 * @param slideNumber - 1-based slide number for the new slide.
 * @param idGenerator - Optional id override forwarded to {@link makeSlideId}.
 */
export function createBlankSlide(slideNumber: number, idGenerator?: () => string): PptxSlide {
	return {
		id: makeSlideId(idGenerator),
		rId: '',
		slideNumber,
		elements: [],
	};
}
