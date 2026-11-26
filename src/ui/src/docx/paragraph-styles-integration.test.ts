// @vitest-environment jsdom
import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import type { DocumentModel, ParagraphStyleCatalog } from 'ooxml-core/docx';
import { loadDocx } from 'ooxml-core/docx';
import { modelToDoc } from './model-adapter';
import { readParagraphFormat } from './paragraph-format';
import { syncStylePicker, paragraphStylesPlugin } from './paragraph-styles';
import { schema } from './schema';
import { paragraphAt } from './test-support';

const catalog: ParagraphStyleCatalog = {
	docDefaults: {},
	styles: {
		Body: { id: 'Body', name: 'Body', isDefault: true, formatting: { align: 'left' } },
		Centered: { id: 'Centered', name: 'Centered', formatting: { align: 'center' } },
	},
	warnings: [],
};

function model(): DocumentModel {
	return {
		blocks: [{ type: 'paragraph', id: 'p1', style: 'Body', runs: [{ text: 'Text' }] }],
		paragraphStyles: catalog,
		page: {
			width: 816,
			height: 1056,
			marginTop: 96,
			marginRight: 96,
			marginBottom: 96,
			marginLeft: 96,
		},
		warnings: [],
	};
}

describe('paragraph style editor integration', () => {
	it('renders inherited table spacing like the direct formatting control (#353)', async () => {
		const spacing = async (name: string) => {
			const path = `../core/docx/layout/fixtures/word-layout-reports/issue-353-${name}.docx`;
			const { model } = await loadDocx(await readFile(path));
			const doc = modelToDoc(model);
			let firstCellParagraph = 0;
			doc.descendants((node, pos, parent) => {
				if (
					!firstCellParagraph &&
					node.type.name === 'paragraph' &&
					parent?.type.name === 'tableCell'
				)
					firstCellParagraph = pos + 1;
			});
			const view = new EditorView(document.createElement('div'), {
				state: EditorState.create({
					doc,
					selection: TextSelection.create(doc, firstCellParagraph),
					plugins: [paragraphStylesPlugin(() => model)],
				}),
			});
			try {
				const read = (selector: string) =>
					Array.from(view.dom.querySelectorAll<HTMLParagraphElement>(selector), ({ style }) => [
						style.marginBottom,
						style.lineHeight,
					]);
				const { afterPt, lineRule } = readParagraphFormat(view.state, model);
				return { cells: read('td p'), body: read('.ProseMirror > p'), afterPt, lineRule };
			} finally {
				view.destroy();
			}
		};
		const control = await spacing('table-direct-spacing-control');
		expect(control.cells).toHaveLength(24);
		expect(control.cells[0]).toEqual(['0px', '1.2']);
		expect(control.body).toHaveLength(1);
		expect(control).toMatchObject({ afterPt: 0, lineRule: 'single' });
		expect(await spacing('table-style-spacing')).toEqual(control);
	});

	it('updates a conditional paragraph style when the editing view changes the table look', async () => {
		const path =
			'../core/docx/layout/fixtures/word-layout-reports/issue-353-table-style-precedence.docx';
		const { model } = await loadDocx(await readFile(path));
		const doc = modelToDoc(model);
		const view = new EditorView(document.createElement('div'), {
			state: EditorState.create({ doc, plugins: [paragraphStylesPlugin(() => model)] }),
		});
		try {
			const paragraph = () => view.dom.querySelector<HTMLParagraphElement>('td p');
			expect(paragraph()?.style.marginTop).toBe('5.33333px');
			expect(paragraph()?.style.marginBottom).toBe('16px');
			expect(paragraph()?.style.textAlign).toBe('center');
			const tablePos = doc.firstChild?.nodeSize ?? 0;
			view.dispatch(view.state.tr.setNodeAttribute(tablePos, 'tableLook', { firstRow: false }));
			expect(paragraph()?.style.marginTop).toBe('');
			expect(paragraph()?.style.marginBottom).toBe('16px');
		} finally {
			view.destroy();
		}
	});

	it('applies a chosen paragraph style at a collapsed text cursor', () => {
		let documentModel = model();
		const doc = schema.node('doc', null, [
			schema.node('paragraph', { id: 'p1', style: 'Body' }, [schema.text('Text')]),
		]);
		const state = EditorState.create({
			doc,
			selection: TextSelection.create(doc, 2),
			plugins: [paragraphStylesPlugin(() => documentModel)],
		});
		const host = document.createElement('div');
		document.body.append(host);
		const view = new EditorView(host, {
			state,
			dispatchTransaction(transaction) {
				const next = view.state.apply(transaction);
				view.updateState(next);
				documentModel = {
					...documentModel,
					blocks: [{ ...documentModel.blocks[0], style: next.doc.firstChild?.attrs.style }],
				} as DocumentModel;
			},
		});
		const toolbar = document.createElement('div');
		toolbar.innerHTML = '<div id="dve-panel-home"></div>';
		syncStylePicker(toolbar, view, documentModel, 'en');
		const select = toolbar.querySelector<HTMLSelectElement>('[data-paragraph-styles]')!;
		select.value = 'Centered';
		select.dispatchEvent(new Event('change', { bubbles: true }));
		expect(view.state.doc.firstChild?.attrs.style).toBe('Centered');
		expect(paragraphAt(documentModel.blocks, 0).style).toBe('Centered');
		expect(view.dom.querySelector('p')?.style.textAlign).toBe('center');
		view.destroy();
		host.remove();
	});
});
