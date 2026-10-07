import type { DocumentModel } from '../index';
import { adaptDocumentModel } from './adapter';
import { layoutSections } from './page-flow';
import { FLOAT_WRAP_NOTE, positionFloats } from './floats';
import { sameExclusions, wrapExclusions } from './wrap';
import type { LayoutDocumentInput } from './input';
import type { LayoutResult } from './result';
import type { TextMeasurer } from './measure';

/** Paginates an engine-native `LayoutDocumentInput` (bypasses the DocumentModel adapter). */
export function layoutDocument(input: LayoutDocumentInput, measurer: TextMeasurer): LayoutResult {
	let result = layoutSections(input, measurer);
	if (!positionFloats(input, result.pages)) return result;
	// Wrapped pictures take room from the text, which can move the pictures' anchors: lay out
	// again around them until their positions settle (usually one extra pass).
	let exclusions = wrapExclusions(result.pages);
	for (let pass = 0; exclusions.size && pass < 3; pass++) {
		result = layoutSections(input, measurer, exclusions);
		positionFloats(input, result.pages);
		const next = wrapExclusions(result.pages);
		if (sameExclusions(next, exclusions)) break;
		exclusions = next;
	}
	return { ...result, approximations: [...new Set([...result.approximations, FLOAT_WRAP_NOTE])] };
}

/** Paginates a `docx-core` `DocumentModel` directly: `adaptDocumentModel` + `layoutDocument`. */
export function layoutDocumentModel(model: DocumentModel, measurer: TextMeasurer): LayoutResult {
	const approximations: string[] = [];
	const input = adaptDocumentModel(model, (message) => approximations.push(message));
	const result = layoutDocument(input, measurer);
	return {
		pages: result.pages,
		approximations: [...new Set([...approximations, ...result.approximations])],
	};
}
