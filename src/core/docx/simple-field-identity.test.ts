import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { Schema } from 'prosemirror-model';
import { EditorState, TextSelection } from 'prosemirror-state';
import { loadDocx } from './parse';
import type { Paragraph } from './model';
import { runToInlineNodes, inlineNodeRun } from './ui/run-adapter';
import { markSpecs } from './ui/schema-marks';
import { fieldResultRanges } from './ui/field-results';
import { commentSelectionRange } from './ui/comment-selection';
import { clearDirectRunFormatting } from './ui/run-format-command';

const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+' },
		paragraph: { content: 'inline*' },
		text: { group: 'inline' },
	},
	marks: markSpecs,
});
it('keeps adjacent identical imported simple fields independent through formatting and export', async () => {
	const loaded = await loadDocx(
		new Uint8Array(
			await readFile(
				new URL('./__fixtures__/field-comments/adjacent-source.docx', import.meta.url),
			),
		),
	);
	const paragraph = loaded.model.blocks[0] as Paragraph;
	const fields = paragraph.runs.filter((run) => run.field);
	expect(fields.map((run) => run.fieldInstanceId)).toEqual([
		'p0:simple-field-0',
		'p0:simple-field-1',
	]);
	const doc = schema.node(
		'doc',
		null,
		schema.node(
			'paragraph',
			null,
			paragraph.runs.flatMap((run) => runToInlineNodes({ ...run, bold: true }, schema)),
		),
	);
	expect(fieldResultRanges(doc).map(({ from, to }) => ({ from, to }))).toEqual([
		{ from: 7, to: 12 },
		{ from: 12, to: 17 },
	]);
	expect(commentSelectionRange(doc, 8, 9)).toEqual({ from: 7, to: 12 });
	let state = EditorState.create({
		doc,
		selection: TextSelection.create(doc, 1, doc.content.size - 1),
	});
	clearDirectRunFormatting(state, (tr) => {
		state = state.apply(tr);
	});
	const runs = Array.from({ length: state.doc.firstChild!.childCount }, (_, index) =>
		inlineNodeRun(state.doc.firstChild!.child(index))!,
	);
	expect(runs.filter((run) => run.field).map((run) => run.fieldInstanceId)).toEqual(
		fields.map((run) => run.fieldInstanceId),
	);
	const bytes = await loaded.save({ ...loaded.model, blocks: [{ ...paragraph, runs }] });
	const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
	expect(xml.match(/<w:fldSimple\b/g)).toHaveLength(2);
	const reloaded = await loadDocx(bytes);
	expect(
		(reloaded.model.blocks[0] as Paragraph).runs.filter((run) => run.field).map((run) => run.text),
	).toEqual(['ABCDE', 'ABCDE']);
});
