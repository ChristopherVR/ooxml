// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { createDocument, type Paragraph, type PictureWatermarkSpec } from 'ooxml-core/docx';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { modelToDoc, docToModel } from './model-adapter';
import { currentWatermark, withWatermark, type AnyWatermark } from './watermark-commands';

const picture: PictureWatermarkSpec = {
	kind: 'picture',
	partName: 'word/media/seal.png',
	contentType: 'image/png',
	widthPt: 300,
	heightPt: 150,
	washout: true,
};

describe('picture watermark in a header story', () => {
	const text: AnyWatermark = {
		text: 'DRAFT',
		color: '#c0c0c0',
		semitransparent: true,
		layout: 'diagonal',
	};
	it.each([
		['picture', picture],
		['text', text],
	] as const)('keeps a %s watermark through typing in the header editor', (_kind, spec) => {
		const withMark = withWatermark(createDocument(), spec, () => 'h1');
		const blocks = withMark.sections![0]!.headers!.default!.blocks;
		let model = { ...createDocument(), blocks: structuredClone(blocks) };
		const view = new EditorView(document.createElement('div'), {
			state: EditorState.create({ doc: modelToDoc(model) }),
			dispatchTransaction(tr) {
				view.updateState(view.state.apply(tr));
				model = docToModel(view.state.doc, model);
			},
		});
		view.dispatch(view.state.tr.insertText('Head', view.state.doc.content.size - 1));
		const runs = (model.blocks[0] as Paragraph).runs;
		expect(runs.map((run) => run.text).join('')).toBe('Head');
		const next = structuredClone(withMark);
		next.sections![0]!.headers!.default!.blocks = model.blocks;
		expect(currentWatermark(next)).toEqual(spec);
		view.destroy();
	});
});
