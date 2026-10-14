// @vitest-environment jsdom
import { describe, expect, it, vi } from 'vitest';
import { EditorState, TextSelection } from 'prosemirror-state';
import { createDocument, type DocumentModel } from 'ooxml-core/docx';
import { countWords, documentWordCount } from 'ooxml-core/docx/ui';
import { createFakeMeasurer, layoutDocument } from 'ooxml-core/docx/layout';
import { docToModel, modelToDoc } from './model-adapter';
import { paragraphStylesPlugin } from './paragraph-styles';
import { runStylesPlugin } from './run-styles';
import { renderPrintLayout } from './print-layout';
import { createCanvasMeasurer } from './canvas-measurer';
import { DecorationSet } from 'prosemirror-view';
import { PrintEditingIndex } from './print-editing-index';

function model(): DocumentModel {
	return {
		...createDocument(),
		blocks: [
			{ type: 'paragraph', id: 'a', runs: [{ text: 'first paragraph' }] },
			{ type: 'paragraph', id: 'b', runs: [{ text: 'second paragraph' }], style: 'Body' },
		],
		paragraphStyles: {
			styles: { Body: { id: 'Body', name: 'Body', formatting: { align: 'right' } } },
			docDefaults: {},
			warnings: [],
		},
		characterStyles: { docDefaults: { bold: true }, styles: {}, warnings: [] },
	};
}

describe('Word performance regression coverage', () => {
	it('reuses Unicode counts for unchanged paragraphs across typing and undo', () => {
		const source = model();
		const doc = modelToDoc(source);
		const state = EditorState.create({ doc });
		const untouched = vi.spyOn(doc.child(1), 'textBetween');
		try {
			expect(documentWordCount(doc, 'en')).toBe(4);
			untouched.mockClear();
			const typed = state.apply(state.tr.insertText('new ', 1));
			expect(documentWordCount(typed.doc, 'en')).toBe(5);
			expect(documentWordCount(doc, 'en')).toBe(4);
			expect(untouched).not.toHaveBeenCalled();
		} finally {
			untouched.mockRestore();
		}
	});

	it.each(['en', 'ar', 'ja', 'not_a_tag'])('matches whole-document counts for %s', (locale) => {
		const source = model();
		source.blocks = ['Hello, world!', 'مرحبا بالعالم', 'これは日本語です', 'cafe\u0301 👩🏽‍💻'].map(
			(text, i) => ({
				type: 'paragraph',
				id: String(i),
				runs: [{ text }],
			}),
		);
		const doc = modelToDoc(source);
		expect(documentWordCount(doc, locale)).toBe(
			countWords(doc.textBetween(0, doc.content.size, ' '), locale),
		);
	});

	it('does not rebuild paragraph or run decorations for caret-only transactions', () => {
		const source = model();
		const state = EditorState.create({ doc: modelToDoc(source) });
		const selected = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 4)));
		for (const plugin of [paragraphStylesPlugin(() => source), runStylesPlugin(() => source)]) {
			const initial = plugin.props.decorations!.call(plugin, state);
			expect(plugin.props.decorations!.call(plugin, selected)).toBe(initial);
		}
	});

	it('shifts cached styled runs after typing before them and invalidates changed catalogs', () => {
		let source = model();
		const state = EditorState.create({ doc: modelToDoc(source) });
		const plugin = runStylesPlugin(() => source);
		const initial = plugin.props.decorations!.call(plugin, state)!;
		expect((initial as DecorationSet).find()).toHaveLength(2);
		const typed = state.apply(state.tr.insertText('new ', 1));
		source = docToModel(typed.doc, source, state.doc);
		const reused = plugin.props.decorations!.call(plugin, typed);
		const fresh = runStylesPlugin(() => source);
		expect(reused).toEqual(fresh.props.decorations!.call(fresh, typed));
		source = { ...source, characterStyles: { docDefaults: {}, styles: {}, warnings: [] } };
		expect(plugin.props.decorations!.call(plugin, typed)).not.toEqual(reused);
	});

	it('matches uncached model conversion while preserving unchanged paragraph identities', () => {
		const source = model();
		const state = EditorState.create({ doc: modelToDoc(source) });
		const typed = state.apply(state.tr.insertText('new ', 1));
		const converted = docToModel(typed.doc, source, state.doc);
		expect(converted).toEqual(docToModel(typed.doc, source));
		expect(converted.blocks[1]).toBe(source.blocks[1]);
		expect(converted.blocks[0]).not.toBe(source.blocks[0]);
	});

	it('measures a font once for ascent and height, but remeasures a different size', () => {
		const measureText = vi.fn(() => ({
			width: 20,
			fontBoundingBoxAscent: 12,
			fontBoundingBoxDescent: 4,
		}));
		const context = vi
			.spyOn(HTMLCanvasElement.prototype, 'getContext')
			.mockReturnValue({ measureText } as unknown as CanvasRenderingContext2D);
		try {
			const measurer = createCanvasMeasurer();
			const font = { family: 'Arial', sizePx: 16 };
			for (let i = 0; i < 100; i++) {
				expect(measurer.ascentOf!(font)).toBe(12);
				expect(measurer.lineHeightOf(font)).toBe(16);
			}
			expect(measureText).toHaveBeenCalledTimes(1);
			measurer.lineHeightOf({ ...font, sizePx: 20 });
			expect(measureText).toHaveBeenCalledTimes(2);
		} finally {
			context.mockRestore();
		}
	});

	it('never reuses an ambiguous paragraph identity in a document with duplicate ids', () => {
		const source = model();
		source.blocks[1]!.id = source.blocks[0]!.id;
		const doc = modelToDoc(source);
		const typed = EditorState.create({ doc }).tr.insertText('x', 1).doc;
		expect(docToModel(typed, source, doc)).toEqual(docToModel(typed, source));
	});

	it('bounds page DOM, keeps snapshots and printing complete, and reindexes mounted pages', () => {
		const result = layoutDocument(
			{
				sections: [
					{
						page: {
							widthPx: 300,
							heightPx: 200,
							marginTopPx: 10,
							marginRightPx: 10,
							marginBottomPx: 10,
							marginLeftPx: 10,
						},
						blocks: Array.from({ length: 60 }, (_, i) => ({
							kind: 'paragraph',
							id: String(i),
							runs: [{ text: `Page ${i}` }],
							pageBreakBefore: true,
						})),
					},
				],
			},
			createFakeMeasurer(),
		);
		const handle = renderPrintLayout(result, undefined, { virtual: true });
		const sheets = [...handle.element.children];
		expect(sheets).toHaveLength(60);
		expect(sheets.filter((sheet) => sheet.childElementCount)).toHaveLength(3);
		const root = document.createElement('div');
		root.append(handle.element);
		const source = model();
		source.blocks = Array.from({ length: 60 }, (_, i) => ({
			type: 'paragraph',
			id: String(i),
			runs: [{ text: `Page ${i}` }],
		}));
		const doc = modelToDoc(source);
		const index = new PrintEditingIndex(root);
		index.update(doc);
		expect(index.block('30')).toHaveLength(0);
		handle.showPages(30, 59);
		index.update(doc);
		expect(index.block('30')).toHaveLength(1);
		expect(index.block('0')).toHaveLength(0);
		expect(sheets.filter((sheet) => sheet.childElementCount)).toHaveLength(6);
		expect(handle.pageSnapshot(10)?.textContent).toContain('Page 10');
		expect(sheets[10]!.childElementCount).toBe(0);
		handle.renderAll();
		expect(sheets.every((sheet) => sheet.childElementCount)).toBe(true);
		handle.showPages(30);
		expect(sheets.filter((sheet) => sheet.childElementCount)).toHaveLength(5);
	});
});
