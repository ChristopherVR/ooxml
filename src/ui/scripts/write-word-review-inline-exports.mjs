// Run with Bun after building core: write synthetic editor exports for native Word checks.
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { EditorState } from 'prosemirror-state';
import { loadDocx } from 'ooxml-core/docx';
import { acceptAllChanges, rejectAllChanges } from 'ooxml-core/docx/ui';
import { modelToDoc, docToModel } from '../src/docx/model-adapter.ts';

const output = process.argv[2];
if (!output) throw new Error('Provide an output directory for synthetic exports');
await mkdir(output, { recursive: true });
const stories = process.argv[3] === '--stories';
for (const name of stories
	? ['all-stories']
	: [
			'picture-insert',
			'picture-delete',
			'note-insert',
			'note-delete',
			'break-delete',
			'break-insert',
		]) {
	for (const mode of ['accept', 'reject']) {
		const loaded = await loadDocx(
			await readFile(
				new URL(
					`../../core/docx/__fixtures__/${stories ? 'review-stories' : 'review-inline'}/${name}-tracked.docx`,
					import.meta.url,
				),
			),
		);
		const view = {
			state: EditorState.create({ doc: modelToDoc(loaded.model) }),
			editable: true,
			dispatch(tr) {
				this.state = this.state.apply(tr);
			},
			focus() {},
		};
		if (!(mode === 'accept' ? acceptAllChanges : rejectAllChanges)(view))
			throw new Error(`Expected ${name} to have a pending revision`);
		const model = docToModel(view.state.doc, loaded.model);
		try {
			await writeFile(resolve(output, `${name}-${mode}.docx`), await loaded.save(model));
		} catch (error) {
			throw new Error(`Cannot export ${name}-${mode}`, { cause: error });
		}
	}
}
console.log(resolve(output));
