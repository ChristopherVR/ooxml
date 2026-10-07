// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { createDocument, loadDocx, saveDocx } from 'ooxml-core/docx';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { addComment } from './comment-commands';
import { docToModel, modelToDoc } from './model-adapter';

describe('comments across paragraphs', () => {
	it('exports and reloads both independent overlapping anchors', async () => {
		const model = createDocument();
		model.blocks = [{ type: 'paragraph', id: 'a', runs: [{ text: 'Alpha beta' }] }];
		const doc = modelToDoc(model);
		const view = new EditorView(document.createElement('div'), {
			state: EditorState.create({ doc, selection: TextSelection.create(doc, 1, 7) }),
		});
		try {
			const first = addComment(view, 'Ada', 'First', () => 'c1')!;
			view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 4, 11)));
			const second = addComment(view, 'Grace', 'Second', () => 'c2')!;
			const next = docToModel(view.state.doc, model);
			const { model: loaded } = await loadDocx(
				await saveDocx({ ...next, comments: [first, second] }),
			);
			expect(loaded.comments?.map((comment) => comment.author)).toEqual(['Ada', 'Grace']);
			const paragraph = loaded.blocks[0]!;
			if (paragraph.type !== 'paragraph') throw new Error('Expected a paragraph');
			expect(
				paragraph.runs
					.filter((run) => run.commentIds?.length === 2)
					.map((run) => run.text)
					.join(''),
			).toBe('ha ');
		} finally {
			view.destroy();
		}
	});
	it('adds one comment over a selection spanning paragraphs and saves a single range', async () => {
		const model = createDocument();
		model.blocks = [
			{ type: 'paragraph', id: 'a', runs: [{ text: 'Alpha one' }] },
			{ type: 'paragraph', id: 'b', runs: [{ text: 'Beta' }] },
			{ type: 'paragraph', id: 'c', runs: [{ text: 'Gamma three' }] },
		];
		const doc = modelToDoc(model);
		// From inside "Alpha one" (after "Alpha ") to inside "Gamma three" (before " three").
		const from = 1 + 'Alpha '.length;
		const to = doc.child(0).nodeSize + doc.child(1).nodeSize + 1 + 'Gamma'.length;
		const view = new EditorView(document.createElement('div'), {
			state: EditorState.create({ doc, selection: TextSelection.create(doc, from, to) }),
		});
		const comment = addComment(view, 'Ada', 'Spans three paragraphs', () => 'c1')!;
		const next = docToModel(view.state.doc, model);
		const saved = await saveDocx({ ...next, comments: [comment] });
		const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string');
		expect(xml.match(/<w:commentRangeStart /g)).toHaveLength(1);
		expect(xml.match(/<w:commentRangeEnd /g)).toHaveLength(1);
		expect(xml).toMatch(/Alpha <\/w:t><\/w:r><w:commentRangeStart w:id="\d+"\/><w:r><w:t>one/);
		expect(xml).toMatch(/Gamma<\/w:t><\/w:r><w:commentRangeEnd w:id="\d+"\/>/);
		// Word requires decimal ids, so the editor's id is renumbered in both parts.
		expect(xml).toContain('<w:commentReference w:id="0"/>');
		const comments = await (
			await JSZip.loadAsync(saved)
		)
			.file('word/comments.xml')!
			.async('string');
		expect(comments).toContain('<w:comment w:id="0" w:author="Ada"');
		expect(comments).not.toContain('c1');
	});
});
