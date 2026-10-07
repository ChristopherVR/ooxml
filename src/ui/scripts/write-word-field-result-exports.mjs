// Run with Bun after building core. Exercise the same result input hooks as the editor.
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { EditorState, TextSelection } from 'prosemirror-state';
import { DOMParser } from 'prosemirror-model';
import { JSDOM } from 'jsdom';
import { loadDocx } from 'ooxml-core/docx';
import { fieldGuardPlugin, fieldResultRanges } from 'ooxml-core/docx/ui';
import { modelToDoc, docToModel } from '../src/docx/model-adapter.ts';

const output = process.argv[2];
if (!output) throw new Error('Provide an output directory');
await mkdir(output, { recursive: true });
const loaded = await loadDocx(
	await readFile(
		new URL('../../core/docx/__fixtures__/field-comments/adjacent-source.docx', import.meta.url),
	),
);
const model = {
	...loaded.model,
	blocks: loaded.model.blocks.map((block) =>
		block.type === 'paragraph'
			? {
					...block,
					runs: block.runs.map((run, index) =>
						run.field && index === 1 ? { ...run, bold: true } : run,
					),
				}
			: block,
	),
};
for (const [kind, offsets] of Object.entries({
	partial: [1, 2],
	whole: [0, 5],
	start: [0, 1],
	end: [4, 5],
})) {
	for (const method of ['typing', 'paste']) {
		const doc = modelToDoc(model);
		const field = fieldResultRanges(doc)[0];
		const guard = fieldGuardPlugin();
		let state = EditorState.create({
			doc,
			selection: TextSelection.create(doc, field.from + offsets[0], field.from + offsets[1]),
			plugins: [guard],
		});
		const view = {
			get state() {
				return state;
			},
			dispatch(tr) {
				state = state.applyTransaction(tr).state;
			},
		};
		if (method === 'typing') {
			if (
				!guard.props.handleTextInput.call(
					guard,
					view,
					state.selection.from,
					state.selection.to,
					'X',
					() => state.tr.insertText('X'),
				)
			)
				throw new Error('Replacement was not handled');
		} else {
			const dom = new JSDOM('<body><b>X</b></body>').window.document.body;
			const slice = DOMParser.fromSchema(doc.type.schema).parseSlice(dom);
			const pasted = guard.props.transformPasted.call(guard, slice, view, false);
			view.dispatch(state.tr.replaceSelection(pasted).setMeta('uiEvent', 'paste'));
		}
		await writeFile(
			resolve(output, `${kind}-${method}.docx`),
			await loaded.save(docToModel(state.doc, model)),
		);
	}
}
console.log(resolve(output));
