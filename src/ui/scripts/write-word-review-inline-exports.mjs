// Run with Bun after building core: write synthetic editor exports for native Word checks.
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { EditorState, TextSelection } from 'prosemirror-state';
import { loadDocx, signedTwips, halfPoints } from 'ooxml-core/docx';
import {
	acceptAllChanges,
	rejectAllChanges,
	trackChangesPlugin,
	createToggleFormat,
	applyRunFormattingPatch,
} from 'ooxml-core/docx/ui';
import { modelToDoc, docToModel } from '../src/docx/model-adapter.ts';

const output = process.argv[2];
if (!output) throw new Error('Provide an output directory for synthetic exports');
await mkdir(output, { recursive: true });
const stories = process.argv[3] === '--stories';
const lineBreak = process.argv[3] === '--record-line-break-formatting';
const advanced = process.argv[3] === '--record-advanced-object-formatting';
const record = advanced || lineBreak || process.argv[3] === '--record-object-formatting';
const objects = record || process.argv[3] === '--object-formatting';
for (const name of advanced
	? ['picture', 'note', 'break', 'field', 'line-break']
	: lineBreak
		? ['line-break']
		: objects
			? ['picture', 'note', 'break', 'field']
			: stories
				? ['all-stories']
				: [
						'picture-insert',
						'picture-delete',
						'note-insert',
						'note-delete',
						'break-delete',
						'break-insert',
					]) {
	for (const mode of record ? ['tracked', 'accept', 'reject'] : ['accept', 'reject']) {
		const loaded = await loadDocx(
			await readFile(
				new URL(
					`../../core/docx/__fixtures__/${advanced ? 'review-advanced-object-formatting' : lineBreak ? 'review-line-break-formatting' : objects ? 'review-object-formatting' : stories ? 'review-stories' : 'review-inline'}/${name}-${record ? 'before' : 'tracked'}.docx`,
					import.meta.url,
				),
			),
		);
		const view = {
			state: EditorState.create({
				doc: modelToDoc(loaded.model),
				plugins: record
					? [
							trackChangesPlugin(
								() => 'Ada',
								() => true,
							),
						]
					: [],
			}),
			editable: true,
			dispatch(tr) {
				this.state = this.state.apply(tr);
			},
			focus() {},
		};
		if (record) {
			let position = -1;
			view.state.doc.descendants((node, pos) => {
				if (position >= 0 || !node.isInline || node.isText) return;
				if (node.type.name === 'fieldMarker' && node.attrs.kind !== 'code') return;
				position = pos;
			});
			if (position < 0) throw new Error(`Missing native ${name} object`);
			view.dispatch(
				view.state.tr.setSelection(TextSelection.create(view.state.doc, position, position + 1)),
			);
			if (advanced)
				applyRunFormattingPatch({
					fontSize: 18,
					color: '#C00000',
					smallCaps: true,
					characterSpacingTwips: signedTwips(30),
					textScalePercent: 150,
					positionHalfPoints: halfPoints(4),
					kerningHalfPoints: halfPoints(24),
				})(view.state, (tr) => view.dispatch(tr));
			else
				createToggleFormat(view.state.schema, () => loaded.model)('bold')(view.state, (tr) =>
					view.dispatch(tr),
				);
		}
		if (mode !== 'tracked' && !(mode === 'accept' ? acceptAllChanges : rejectAllChanges)(view))
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
