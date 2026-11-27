// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Draws the results of computed fields (DOCPROPERTY, AUTONUM, LISTNUM, SYMBOL...) in the layout
// input. Each run keeps its source length, so editor positions still map onto the document: a
// result replaces the text of the field's result runs, and a field Word saves without a result
// draws on its `end` marker, which takes one position like every field marker.
import { documentParagraphs } from '../cell-blocks';
import { computedFields, type ComputedFieldResult } from '../computed-fields';
import { fieldResultSpans } from '../field-result-spans';
import type { DocumentModel, Paragraph } from '../model';
import type { LayoutBlock, LayoutDocumentInput, LayoutParagraph, LayoutRun } from './input';
import { reviewSourceLength } from './review-display';
import { allCapsText, smallCapsText } from '../caps-display';

/** A stretch of a paragraph's source positions and what it draws. */
interface Patch {
	from: number;
	to: number;
	result: ComputedFieldResult | undefined;
	/** The text the model holds there. */
	source: string;
}

/** The source stretches of `paragraph` whose drawing a computed field changes. */
function patches(paragraph: Paragraph, fields: ReturnType<typeof computedFields>): Patch[] {
	if (!paragraph.runs.some((run) => run.field || run.fieldChar)) return [];
	const starts: number[] = [];
	let offset = 0;
	for (const run of paragraph.runs) {
		starts.push(offset);
		offset += reviewSourceLength(run);
	}
	const stretch = (at: number) => ({
		from: starts[at]!,
		to: starts[at]! + reviewSourceLength(paragraph.runs[at]!),
		source: paragraph.runs[at]!.text,
	});
	const found: Patch[] = [];
	for (const field of fieldResultSpans(paragraph.runs, { resultless: true })) {
		if (field.locked) continue;
		const result = fields.result(field, paragraph);
		if (!result) continue;
		if (field.resultless) found.push({ ...stretch(field.to), result });
		else
			field.runs.forEach((at, index) =>
				found.push({ ...stretch(at), result: index === 0 ? result : undefined }),
			);
	}
	return found;
}

/** A layout run drawing `result` (or nothing) in place of its own text, at its source length. */
function patched(run: LayoutRun, patch: Patch, first: boolean): LayoutRun {
	const { result } = patch;
	const sourceLength = run.sourceLength ?? run.text.length;
	const text = first && result ? result.text : '';
	// A run drawn in capitals (`w:caps`) shows its result in capitals too, in French (France)
	// without the accents of the letters the formatting capitalised.
	const capsLanguage =
		run.text === patch.source
			? undefined
			: run.text === patch.source.toUpperCase()
				? ''
				: run.text === allCapsText(patch.source, 'fr-FR')
					? 'fr-FR'
					: undefined;
	return {
		...run,
		text:
			capsLanguage !== undefined
				? allCapsText(text, capsLanguage)
				: // A French (France) small caps run lost its accents the same way.
					run.smallCaps && run.text !== patch.source
					? smallCapsText(text, 'fr-FR')
					: text,
		sourceLength,
		...(first && result?.fontFamily ? { fontFamily: result.fontFamily } : {}),
		...(first && result?.fontSizePt ? { fontSizePt: result.fontSizePt } : {}),
	};
}

function paragraphWithResults(
	paragraph: LayoutParagraph,
	model: Paragraph | undefined,
	fields: ReturnType<typeof computedFields>,
): LayoutParagraph {
	const list = model ? patches(model, fields) : [];
	if (!list.length) return paragraph;
	let offset = 0;
	let changed = false;
	const runs = paragraph.runs.map((run) => {
		if (run.synthetic) return run;
		const from = offset;
		offset += run.sourceLength ?? run.text.length;
		const patch = list.find((item) => from >= item.from && from < item.to);
		// Hidden text (`w:vanish`) draws nothing, whatever its result.
		if (!patch || run.hidden || run.object || (run.text === '' && patch.source !== '')) return run;
		changed = true;
		return patched(run, patch, from === patch.from);
	});
	return changed ? { ...paragraph, runs } : paragraph;
}

/**
 * `input` with the results of `model`'s computed fields drawn in its paragraphs (in tables too).
 * `model` is the one `input` was adapted from. The input itself when nothing changes.
 */
export function withComputedFieldRuns(
	input: LayoutDocumentInput,
	model: DocumentModel,
): LayoutDocumentInput {
	const fields = computedFields(model);
	const paragraphs = new Map(documentParagraphs(model.blocks).map((item) => [item.id, item]));
	const block = (item: LayoutBlock): LayoutBlock => {
		if (item.kind === 'paragraph')
			return paragraphWithResults(item, paragraphs.get(item.id), fields);
		let changed = false;
		const rows = item.rows.map((row) => {
			const cells = row.cells.map((cell) => {
				const next = cell.paragraphs.map((paragraph) => block(paragraph) as LayoutParagraph);
				const blocks = cell.blocks?.map(block);
				if (
					next.every((paragraph, index) => paragraph === cell.paragraphs[index]) &&
					(blocks ?? []).every((child, index) => child === cell.blocks![index])
				)
					return cell;
				changed = true;
				return { ...cell, paragraphs: next, ...(blocks ? { blocks } : {}) };
			});
			return cells.every((cell, index) => cell === row.cells[index]) ? row : { ...row, cells };
		});
		return changed ? { ...item, rows } : item;
	};
	let changed = false;
	const sections = input.sections.map((section) => {
		const blocks = section.blocks.map(block);
		if (blocks.every((item, index) => item === section.blocks[index])) return section;
		changed = true;
		return { ...section, blocks };
	});
	return changed ? { ...input, sections } : input;
}
