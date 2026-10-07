import { describe, expect, it } from 'vitest';
import { Schema } from 'prosemirror-model';
import { EditorState, type Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import {
	acceptRevisionRange,
	collectRevisionRanges,
	goToNextChange,
	rejectRevisionRange,
} from './review-commands.js';

const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+' },
		paragraph: {
			content: 'text*',
			attrs: {
				id: { default: '' },
				align: { default: null },
				formatRevision: { default: null },
				restoredParagraphPropertiesXml: { default: null },
			},
		},
		text: {},
	},
});
function view(editable = true): EditorView {
	const paragraph = (id: string, author: string) =>
		schema.node('paragraph', {
			id,
			align: 'center',
			formatRevision: {
				kind: 'paragraphChange',
				id: 'legacy',
				author,
				previousParagraphPropertiesXml:
					'<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:jc w:val="left"/></w:pPr>',
			},
		});
	const host = {
		state: EditorState.create({
			doc: schema.node('doc', null, [paragraph('a', 'Ada'), paragraph('b', 'Grace')]),
		}),
		editable,
		dispatch(tr: Transaction) {
			this.state = this.state.apply(tr);
		},
		focus() {},
	};
	return host as unknown as EditorView;
}
describe('paragraph review boundaries', () => {
	it('navigates and resolves an empty paragraph without resolving a colliding peer id', () => {
		const editor = view();
		expect(goToNextChange(editor)).toBe(true);
		const range = collectRevisionRanges(editor.state.doc)[0]!;
		expect(range.from).toBe(range.to);
		rejectRevisionRange(editor, range);
		expect(editor.state.doc.firstChild!.attrs.align).toBe('left');
		expect(collectRevisionRanges(editor.state.doc).map((item) => item.author)).toEqual(['Grace']);
		rejectRevisionRange(editor, range);
		expect(collectRevisionRanges(editor.state.doc).map((item) => item.author)).toEqual(['Grace']);
	});
	it('blocks paragraph acceptance and rejection in a read-only view', () => {
		const editor = view(false);
		const before = editor.state.doc;
		const range = collectRevisionRanges(before)[0]!;
		acceptRevisionRange(editor, range);
		rejectRevisionRange(editor, range);
		expect(editor.state.doc).toBe(before);
	});
});
