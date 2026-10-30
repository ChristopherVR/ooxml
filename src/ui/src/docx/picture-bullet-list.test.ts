// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { createDocument } from 'ooxml-core/docx';
import { schema } from './schema';
import { applyPictureBulletList } from './list-commands';

describe('picture bullet list', () => {
	it('adds a numPicBullet entry and numbers the selected paragraph with it', () => {
		const doc = schema.nodes.doc.create(null, [
			schema.node('paragraph', { id: 'p0' }, schema.text('Item')),
		]);
		const view = new EditorView(document.createElement('div'), {
			state: EditorState.create({ doc }),
		});
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1)));
		const model = createDocument();
		applyPictureBulletList(view, model, {
			partName: 'word/media/bullet.png',
			contentType: 'image/png',
		});
		const bullet = model.numberingCatalog?.pictureBullets?.['0'];
		expect(bullet?.partName).toBe('word/media/bullet.png');
		const numId = String(view.state.doc.child(0).attrs.numId);
		const abstractId = model.numberingCatalog?.nums[numId]?.abstractNumId ?? '';
		expect(model.numberingCatalog?.abstractNums[abstractId]?.levels[0]?.pictureBulletId).toBe('0');
	});
});
