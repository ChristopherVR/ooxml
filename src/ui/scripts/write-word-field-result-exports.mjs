// Run with Bun after building core. Exercise the same result input hooks as the editor.
import { mkdir, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { EditorState, TextSelection } from 'prosemirror-state';
import { DOMParser } from 'prosemirror-model';
import { JSDOM } from 'jsdom';
import { fieldGuardPlugin, fieldResultRanges, trackChangesPlugin } from 'ooxml-core/docx/ui';
import { modelToDoc, docToModel } from '../src/docx/model-adapter.ts';
import { fieldExportSource } from './word-field-export-source.mjs';

if (!process.argv[2]) throw new Error('Provide an output directory');
const output = resolve(process.argv[2]);
await mkdir(output, { recursive: true });
const { loaded, model } = await fieldExportSource();
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

const emptyDoc = modelToDoc(model);
const emptyField = fieldResultRanges(emptyDoc)[0];
const emptyGuard = fieldGuardPlugin();
let emptyState = EditorState.create({
	doc: emptyDoc,
	selection: TextSelection.create(emptyDoc, emptyField.from, emptyField.to),
	plugins: [emptyGuard],
});
const emptyView = {
	get state() {
		return emptyState;
	},
	dispatch(tr) {
		emptyState = emptyState.applyTransaction(tr).state;
	},
};
if (!emptyGuard.props.handleKeyDown.call(emptyGuard, emptyView, { key: 'Backspace' }))
	throw new Error('Empty result deletion was not handled');
await writeFile(
	resolve(output, 'empty-delete.docx'),
	await loaded.save(docToModel(emptyState.doc, model)),
);

emptyState = EditorState.create({
	doc: emptyDoc,
	selection: TextSelection.create(emptyDoc, emptyField.from, emptyField.to),
	plugins: [
		emptyGuard,
		trackChangesPlugin(
			() => 'Ada',
			() => true,
		),
	],
});
emptyGuard.props.handleKeyDown.call(emptyGuard, emptyView, { key: 'Backspace' });
await writeFile(
	resolve(output, 'tracked-empty-delete.docx'),
	await loaded.save(docToModel(emptyState.doc, model)),
);
