// Word writes a text form field that holds no text as five en spaces (U+2002), which keep the
// shaded field visible. The editor treats such a blank as empty, as Word does when the field is
// filled in: the first text typed into it takes the en spaces' place, and they never count
// against the field's maximum length.
import type { Node } from 'prosemirror-model';
import type { Mappable } from 'prosemirror-transform';
import type { FormTextResult } from './form-field-commands';

const EN_SPACE = '\u2002';

type Range = { from: number; to: number };

/** Whether a text field's result is Word's blank: en spaces only. */
export function isBlankFormText(doc: Node, result: FormTextResult): boolean {
	// A paragraph break or an object inside the result is another character.
	const text = doc.textBetween(result.from, result.to, '\n', '\uFFFC');
	return text.length > 0 && [...text].every((character) => character === EN_SPACE);
}

/**
 * The en spaces of blank results that an edit typed other text into, where they are after the
 * edit, one range per character: they go, and the text typed stays. `before` and `after` are the
 * documents before and after the edit, each with its text field results, and `mapping` maps the
 * one to the other.
 */
export function replacedBlanks(
	before: Node,
	beforeResults: readonly FormTextResult[],
	after: Node,
	afterResults: readonly FormTextResult[],
	mapping: Mappable,
): Range[] {
	const left: Range[] = [];
	for (const result of beforeResults) {
		if (!isBlankFormText(before, result)) continue;
		for (let at = result.from; at < result.to; at++) {
			// An en space that is still there sits between its two edges, mapped away from text
			// typed beside it.
			const from = mapping.map(at, 1);
			const to = mapping.map(at + 1, -1);
			if (to - from === 1 && after.textBetween(from, to) === EN_SPACE) left.push({ from, to });
		}
	}
	if (!left.length) return left;
	const filled = afterResults.filter((result) => !isBlankFormText(after, result));
	return left.filter((space) =>
		filled.some((result) => space.from >= result.from && space.to <= result.to),
	);
}
