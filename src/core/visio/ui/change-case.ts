import type { VisioPage, VisioShape } from '../model';
import type { VisioTextRangesEdit } from '../edit-text-range-commands';
import { changeTextCaseRanges, TEXT_CASE_MODES, type TextCaseMode } from '../../text/change-case';

export type { TextCaseMode } from '../../text/change-case';

function findShape(shapes: readonly VisioShape[], id: string, depth = 0): VisioShape | undefined {
	if (depth > 64) return undefined;
	for (const shape of shapes) {
		if (shape.id === id) return shape;
		const child = findShape(shape.children, id, depth + 1);
		if (child) return child;
	}
	return undefined;
}

/** A shape whose displayed text Change Case can rewrite; source admission stays authoritative. */
export function visioChangeCaseShape(page: VisioPage, shapeId: string): VisioShape | undefined {
	const shape = findShape(page.shapes, shapeId);
	return shape && !shape.hidden && shape.text.plainText.length > 0 ? shape : undefined;
}

/**
 * Change Case as a source-preserving range edit. Ranges never cross a run (text node) or line
 * boundary, so each character keeps its own formatting; returns undefined when nothing changes.
 */
export function visioChangeCaseCommand(
	page: VisioPage,
	shapeId: string,
	mode: TextCaseMode,
): VisioTextRangesEdit | undefined {
	if (!TEXT_CASE_MODES.includes(mode)) return undefined;
	const shape = visioChangeCaseShape(page, shapeId);
	if (!shape) return undefined;
	const text = shape.text.plainText;
	const boundaries: number[] = [];
	let offset = 0;
	for (const run of shape.text.runs) {
		boundaries.push(offset);
		offset += run.text.length;
	}
	const ranges = changeTextCaseRanges(text, mode, boundaries);
	if (!ranges.length) return undefined;
	return {
		type: 'replace-text-ranges',
		pageId: page.id,
		shapeId: shape.id,
		expectedText: text,
		ranges,
	};
}
