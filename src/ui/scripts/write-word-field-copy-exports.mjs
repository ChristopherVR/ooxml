// Run with Bun after building core. Transfers use the shared clipboard hooks and UI model adapter.
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { EditorState, TextSelection } from 'prosemirror-state';
import { loadDocx } from 'ooxml-core/docx';
import { fieldGuardPlugin, fieldResultRanges } from 'ooxml-core/docx/ui';
import { modelToDoc, docToModel } from '../src/docx/model-adapter.ts';

const output = process.argv[2];
if (!output) throw new Error('Provide an output directory');
await mkdir(output, { recursive: true });
const loaded = await loadDocx(
	await readFile(
		new URL('../../core/docx/__fixtures__/field-range-copy/result.docx', import.meta.url),
	),
);
const model = {
	...loaded.model,
	blocks: loaded.model.blocks.map((block) =>
		block.type === 'paragraph'
			? {
					...block,
					runs: block.runs.filter((run) => run.text !== 'ABCDE' || run.field),
				}
			: block,
	),
};
for (const kind of ['whole', 'result', 'partial']) {
	const doc = modelToDoc(model);
	const field = fieldResultRanges(doc)[0];
	if (!field) throw new Error('Expected a field result');
	const from = kind === 'whole' ? field.from - 3 : kind === 'partial' ? field.from + 1 : field.from;
	const to = kind === 'whole' ? field.to + 1 : kind === 'partial' ? field.from + 2 : field.to;
	const guard = fieldGuardPlugin();
	let state = EditorState.create({
		doc,
		selection: TextSelection.create(doc, from, to),
		plugins: [guard],
	});
	let slice = guard.props.transformCopied.call(guard, state.selection.content(), undefined);
	slice = guard.props.transformPasted.call(guard, slice, undefined, false);
	state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, field.to + 1)));
	state = state.apply(state.tr.replaceSelection(slice).setMeta('uiEvent', 'paste'));
	const next = docToModel(state.doc, model);
	await writeFile(resolve(output, `${kind}.docx`), await loaded.save(next));
}
console.log(resolve(output));
