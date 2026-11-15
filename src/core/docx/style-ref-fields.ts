// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Brings the cached results of STYLEREF fields up to date in a story's paragraphs.
import { withFieldResults } from './field-result-spans';
import { withStoryBlocks } from './field-stories';
import type { Block, DocumentModel, Paragraph } from './model';
import {
	parseStyleReference,
	styleRefSource,
	type RunSpan,
	type StyleRefSource,
	type StyleReference,
} from './style-ref';

/** Gives a STYLEREF field its result, or undefined to keep the one it has. */
export type StyleRefEvaluator = (
	reference: StyleReference,
	paragraph: Paragraph,
	/** The field's own runs in `paragraph`. */
	field: RunSpan,
) => string | undefined;

/**
 * The blocks with every unlocked STYLEREF result that `evaluate` answers set to its answer.
 * Unchanged paragraphs and tables are returned as they are, so the writer keeps their XML.
 */
export function withStyleRefResults(
	blocks: readonly Block[],
	evaluate: StyleRefEvaluator,
): Block[] {
	return withFieldResults(blocks, (field, paragraph) => {
		const reference = parseStyleReference(field.instr);
		const text = reference && evaluate(reference, paragraph, field);
		return text === undefined ? undefined : { text, charFormat: reference?.charFormat === true };
	});
}

/**
 * The result a header or footer STYLEREF shows on every page, when that does not depend on the
 * page: every text in the document it could pick up reads the same (typically the style is used
 * once, for a title or a document number). Undefined otherwise, or when nothing is in the style.
 */
function pageIndependentResult(
	source: StyleRefSource,
	reference: StyleReference,
): string | undefined {
	const target = source.target(reference.style);
	if (!target) return undefined;
	let result: string | undefined;
	for (let index = 0; index < source.paragraphs.length; index++)
		for (const text of source.results(index, reference, target)) {
			if (text === undefined || (result !== undefined && text !== result)) return undefined;
			result = text;
		}
	return result;
}

/**
 * The result of a STYLEREF in the body: the nearest text in the style before the field, else the
 * nearest after it (ECMA-376 Part 1, 17.16.5; Word searches back from the field, then on from
 * it). Text in a character style in the field's own paragraph is nearest, the field's own runs
 * (`field`) left out. A paragraph style's match is a whole paragraph: the field's own is the
 * nearest for its number (`\n`), which Word shows when a paragraph repeats its own number in its
 * text, but not for its text, which holds the field's own result. Undefined when the paragraph
 * is not in the main document or nothing is in the style.
 */
export function bodyStyleRefResult(
	source: StyleRefSource,
	reference: StyleReference,
	paragraph: Paragraph,
	field?: RunSpan,
): string | undefined {
	const target = source.target(reference.style);
	const at = source.index.get(paragraph.id);
	if (!target || at === undefined) return undefined;
	const own =
		field && target.kind === 'character' ? source.portions(at, reference, target, field) : [];
	const earlier = own.filter((portion) => portion.to < field!.from);
	if (earlier.length) return earlier.at(-1)!.text;
	const number =
		target.kind === 'paragraph' && reference.number ? source.results(at, reference, target) : [];
	if (number.length) return number[0];
	for (let index = at - 1; index >= 0; index--) {
		const found = source.results(index, reference, target);
		if (found.length) return found.at(-1);
	}
	const later = own.find((portion) => portion.from > field!.to);
	if (later) return later.text;
	for (let index = at + 1; index < source.paragraphs.length; index++) {
		const found = source.results(index, reference, target);
		if (found.length) return found[0];
	}
	return undefined;
}

/**
 * Brings STYLEREF results up to date. A field in a header or footer takes the text it shows on
 * every page, when that does not depend on the page (`pageIndependentResult`); Word updates such
 * fields as the page is drawn, so a saved header should not keep a stale value. With
 * `body: true`, fields in the body take the nearest text in their style before them, else after
 * them, as Word's Update Field does. Locked fields and fields with switches that are not evaluated
 * keep their result. Returns `model` itself when nothing changed.
 */
export function refreshStyleRefFields(
	model: DocumentModel,
	options: { body?: boolean } = {},
): DocumentModel {
	let source: StyleRefSource | undefined;
	const read = () => (source ??= styleRefSource(model));
	const everywhere = new Map<string, string | undefined>();
	const story: StyleRefEvaluator = (reference) => {
		const key = JSON.stringify(reference);
		if (!everywhere.has(key)) everywhere.set(key, pageIndependentResult(read(), reference));
		return everywhere.get(key);
	};
	return withStoryBlocks(
		model,
		(blocks) => withStyleRefResults(blocks, story),
		options.body
			? (blocks) =>
					withStyleRefResults(blocks, (reference, paragraph, field) =>
						bodyStyleRefResult(read(), reference, paragraph, field),
					)
			: undefined,
	);
}
