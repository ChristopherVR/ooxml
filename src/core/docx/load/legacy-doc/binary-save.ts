// Saving a converted Word 97-2003 document back to the binary .doc format. The binary writer only
// replaces the text of existing main-story paragraphs, so this accepts a model that is unchanged,
// or one whose only change is the text of single-run top-level paragraphs that line up with the
// file's own paragraphs. Everything else is refused; `saveDocx` writes the full model instead.
import {
	readOleDocParagraphs,
	writeOleDocParagraphEdit,
} from '@christophervr/ole2/ole-document-doc-editor';
import type { DocumentModel, Paragraph } from '../../model';
import { LegacyDocError } from './error';

const refuse = (message: string) =>
	new LegacyDocError(`${message} Save the document as .docx (saveDocx) to keep this change.`);

/** The model with the text of top-level paragraph runs and the load warnings left out. */
function shape(model: DocumentModel): string {
	const { warnings: _warnings, ...rest } = model;
	return JSON.stringify({
		...rest,
		blocks: model.blocks.map((block) =>
			block.type === 'paragraph'
				? { ...block, runs: block.runs.map((run) => ({ ...run, text: '' })) }
				: block,
		),
	});
}

const paragraphText = (p: Paragraph) => p.runs.map((run) => run.text).join('');

/** Prepares the binary save for `original`, whose converted model is `model`. */
export function binaryDocSaver(original: Uint8Array, model: DocumentModel) {
	const content = ({ warnings: _warnings, ...rest }: DocumentModel) => JSON.stringify(rest);
	const initial = content(model);
	const initialShape = shape(model);
	const filePara = readOleDocParagraphs(original);
	// Text edits are written by paragraph index, so the model must list exactly the file's own
	// main-story paragraphs, in order and with the same text.
	const aligned =
		!!filePara &&
		model.blocks.length === filePara.length &&
		model.blocks.every((b, i) => b.type === 'paragraph' && paragraphText(b) === filePara[i]);

	return (next: DocumentModel = model): Uint8Array => {
		if (content(next) === initial) return original.slice();
		if (shape(next) !== initialShape)
			throw refuse('Only paragraph text edits can be written to a Word 97-2003 .doc file.');
		if (!aligned || !filePara)
			throw refuse(
				'This .doc has tables, fields or other structures that prevent paragraph text editing.',
			);
		let output = original;
		next.blocks.forEach((block, i) => {
			const text = paragraphText(block as Paragraph);
			if (text === filePara[i]) return;
			if ((block as Paragraph).runs.length !== 1)
				throw refuse(
					'Editing text in a paragraph with mixed formatting is not supported in a .doc file.',
				);
			if (/[\r\n]/u.test(text))
				throw refuse('Adding line breaks in legacy .doc paragraph text is unsupported.');
			const result = writeOleDocParagraphEdit(output, i, text);
			if (result === output)
				throw refuse('This .doc uses structures that prevent safe paragraph text editing.');
			output = result;
		});
		return output;
	};
}
