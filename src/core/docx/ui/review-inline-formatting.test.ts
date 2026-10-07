import { expect, it } from 'vitest';
import { Schema } from 'prosemirror-model';
import { EditorState, type Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { history, undo, redo } from 'prosemirror-history';
import type { TextRun } from '../model';
import { markSpecs } from './schema-marks';
import { imageNodeSpec } from './inline-content-schema';
import { pageBreakNodeSpec, noteReferenceNodeSpec, fieldMarkerNodeSpec } from './break-note-schema';
import { inlineNodeRun, runToInlineNodes } from './run-adapter';
import { collectRevisionRanges, acceptRevisionRange, rejectRevisionRange } from './review-commands';

const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+' },
		paragraph: { content: 'inline*' },
		text: { group: 'inline' },
		image: imageNodeSpec,
		pageBreak: pageBreakNodeSpec,
		noteReference: noteReferenceNodeSpec,
		fieldMarker: fieldMarkerNodeSpec,
	},
	marks: markSpecs,
});
const atoms: TextRun[] = [
	{ text: '', break: 'page' },
	{ text: '', noteReference: { kind: 'endnote', id: '7' } },
	{ text: '', fieldChar: 'begin' },
	{ text: '', fieldCode: ' PAGE ' },
	{
		text: '',
		image: {
			relId: 'rId1',
			partName: 'word/media/image.png',
			contentType: 'image/png',
			widthPx: 20,
			heightPx: 30,
		},
	},
];
const snapshot =
	'<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:b w:val="0"/><w:i/><w:lang w:val="en-US"/><w:noProof/></w:rPr>';
for (const atom of atoms)
	for (const mode of ['accept', 'reject'] as const)
		it(`${mode}s ${Object.keys(atom)[1]} formatting independently of insertion history`, () => {
			const run: TextRun = {
				...atom,
				bold: true,
				language: 'fr-FR',
				fontFamilyComplexScript: 'Arial',
				commentIds: ['comment'],
				link: { href: 'https://example.test/' },
				revision: { id: 'text', kind: 'insert', author: 'Ada' },
				formatRevision: {
					id: 'format',
					kind: 'formatChange',
					author: 'Bob',
					previousRunPropertiesXml: snapshot,
				},
			};
			const host = {
				state: EditorState.create({
					doc: schema.node(
						'doc',
						null,
						schema.node(
							'paragraph',
							null,
							runToInlineNodes(run, schema, () => ({ number: 42, label: 'XLII' })),
						),
					),
					plugins: [history()],
				}),
				editable: true,
				dispatch(tr: Transaction) {
					this.state = this.state.apply(tr);
				},
				focus() {},
			};
			const editor = host as unknown as EditorView;
			const initial = editor.state.doc;
			const format = collectRevisionRanges(initial).find((range) => range.kind === 'formatChange')!;
			expect(format.author).toBe('Bob');
			(mode === 'accept' ? acceptRevisionRange : rejectRevisionRange)(editor, format);
			const node = editor.state.doc.firstChild!.firstChild!;
			const actual = inlineNodeRun(node)!;
			expect(actual).toMatchObject(atom);
			expect(actual.revision).toEqual(run.revision);
			expect(actual.formatRevision).toBeUndefined();
			expect(actual.commentIds).toEqual(['comment']);
			expect(actual.link).toEqual(run.link);
			expect(actual.bold).toBe(mode === 'accept');
			expect(actual.italic).toBe(mode === 'reject' ? true : undefined);
			expect(actual.language).toBe(mode === 'accept' ? 'fr-FR' : 'en-US');
			expect(actual.fontFamilyComplexScript).toBe(mode === 'accept' ? 'Arial' : undefined);
			if (mode === 'reject') expect(actual.sourceRunPropertiesXml).toContain('noProof');
			if (atom.noteReference) expect(node.attrs).toMatchObject({ number: 42, label: 'XLII' });
			expect(collectRevisionRanges(editor.state.doc)).toMatchObject([
				{ kind: 'insert', author: 'Ada', id: 'text' },
			]);
			const resolved = editor.state.doc;
			expect(undo(editor.state, editor.dispatch.bind(editor))).toBe(true);
			expect(editor.state.doc.eq(initial)).toBe(true);
			expect(redo(editor.state, editor.dispatch.bind(editor))).toBe(true);
			expect(editor.state.doc.eq(resolved)).toBe(true);
		});
