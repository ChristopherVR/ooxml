// Run with Bun after building core. Clipboard and deletion use the editor's shared core hooks.
import { mkdir, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { EditorState, TextSelection } from 'prosemirror-state';
import { fieldGuardPlugin, fieldResultRanges, deleteFieldSelection } from 'ooxml-core/docx/ui';
import { modelToDoc, docToModel } from '../src/docx/model-adapter.ts';
import { fieldExportSource } from './word-field-export-source.mjs';
if (!process.argv[2]) throw new Error('Provide an output directory');
const output = resolve(process.argv[2]);
await mkdir(output, { recursive: true });
const { loaded, model } = await fieldExportSource();
for (const kind of ['result', 'partial']) {
	const doc = modelToDoc(model);
	const field = fieldResultRanges(doc)[0];
	const guard = fieldGuardPlugin();
	let state = EditorState.create({
		doc,
		selection: TextSelection.create(
			doc,
			kind === 'result' ? field.from : field.from + 1,
			kind === 'result' ? field.to : field.from + 2,
		),
		plugins: [guard],
	});
	const copied = guard.props.transformCopied.call(guard, state.selection.content(), undefined);
	state = state.applyTransaction(
		(deleteFieldSelection(state) ?? state.tr.deleteSelection()).setMeta('uiEvent', 'cut'),
	).state;
	let destination = fieldResultRanges(state.doc)[0].to;
	if (kind === 'result')
		state.doc.descendants((node, pos) => {
			if (node.attrs.kind === 'end' && pos >= field.from) {
				destination = pos + 1;
				return false;
			}
		});
	state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, destination)));
	const view = { state };
	const pasted = guard.props.transformPasted.call(guard, copied, view, false);
	state = state.applyTransaction(
		state.tr.replaceSelection(pasted).setMeta('uiEvent', 'paste'),
	).state;
	await writeFile(resolve(output, `${kind}.docx`), await loaded.save(docToModel(state.doc, model)));
}
console.log(output);
