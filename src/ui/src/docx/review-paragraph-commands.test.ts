import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { loadDocx, type Paragraph } from 'ooxml-core/docx';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { history, undo, redo } from 'prosemirror-history';
import {
	acceptRevisionRange,
	collectRevisionRanges,
	goToNextChange,
	rejectRevisionRange,
} from 'ooxml-core/docx/ui';
import { trackChangesPlugin } from './track-changes-mode';
import { modelToDoc, docToModel } from './model-adapter';

const views: EditorView[] = [];
afterEach(() => {
	for (const view of views.splice(0)) view.destroy();
	document.body.replaceChildren();
});
for (const name of ['alignment', 'spacing', 'indent', 'multiple'])
	describe(`review native paragraph ${name}`, () => {
		for (const mode of ['accept', 'reject'] as const)
			it(`${mode} preserves text and supports undo/redo with recording enabled`, async () => {
				const loaded = await loadDocx(
					new Uint8Array(
						await readFile(
							resolve(
								'../core/docx/__fixtures__/review-paragraph-formatting',
								`${name}-tracked.docx`,
							),
						),
					),
				);
				const host = document.body.appendChild(document.createElement('div'));
				const view = new EditorView(host, {
					state: EditorState.create({
						doc: modelToDoc(loaded.model),
						plugins: [
							history(),
							trackChangesPlugin(
								() => 'Reviewer',
								() => true,
							),
						],
					}),
				});
				views.push(view);
				const before = view.state.doc;
				const ranges = collectRevisionRanges(before);
				expect(ranges).toHaveLength(1);
				expect(ranges[0]!.kind).toBe('paragraphChange');
				expect(goToNextChange(view)).toBe(true);
				expect(view.state.selection.from).toBe(ranges[0]!.from);
				(mode === 'accept' ? acceptRevisionRange : rejectRevisionRange)(view, ranges[0]!);
				expect(view.state.doc.textContent).toBe('Paragraph formatting');
				expect(collectRevisionRanges(view.state.doc)).toEqual([]);
				const model = docToModel(view.state.doc, loaded.model);
				const actual = model.blocks[0] as Paragraph;
				if (mode === 'reject') {
					const native = await loadDocx(
						new Uint8Array(
							await readFile(
								resolve(
									'../core/docx/__fixtures__/review-paragraph-formatting',
									`${name}-rejected.docx`,
								),
							),
						),
					);
					const { restoredParagraphPropertiesXml: _snapshot, ...properties } = actual;
					expect(properties).toEqual(native.model.blocks[0]);
				} else {
					const { formatRevision: _revision, ...expected } = loaded.model.blocks[0] as Paragraph;
					expect(actual).toEqual(expected);
				}
				const reopened = await loadDocx(await loaded.save(model));
				expect((reopened.model.blocks[0] as Paragraph).formatRevision).toBeUndefined();
				expect(undo(view.state, view.dispatch)).toBe(true);
				expect(view.state.doc.eq(before)).toBe(true);
				expect(redo(view.state, view.dispatch)).toBe(true);
				expect(collectRevisionRanges(view.state.doc)).toEqual([]);
			});
	});
