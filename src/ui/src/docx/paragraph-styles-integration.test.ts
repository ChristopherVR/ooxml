// @vitest-environment jsdom
import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import type { DocumentModel, ParagraphStyleCatalog } from 'ooxml-core/docx';
import { editStyleCatalogs, loadDocx } from 'ooxml-core/docx';
import { readFontFormat } from './font-format';
import { modelToDoc } from './model-adapter';
import { readParagraphFormat } from './paragraph-format';
import { syncStylePicker, paragraphStylesPlugin } from './paragraph-styles';
import { runStylesPlugin } from './run-styles';
import { schema } from './schema';
import { must, paragraphAt } from './test-support';

const tableStyleFixture =
	'../core/docx/layout/fixtures/word-layout-reports/issue-353-table-style-precedence.docx';

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
	it('updates the section break marker and list count when the break gains text (#355)', async () => {
		const path =
			'../core/docx/layout/fixtures/word-layout-reports/issue-355-numbered-section-break-nextpage.docx';
		const { model } = await loadDocx(await readFile(path));
		const doc = modelToDoc(model);
		const view = new EditorView(document.createElement('div'), {
			state: EditorState.create({ doc, plugins: [paragraphStylesPlugin(() => model)] }),
		});
		try {
			const markers = () =>
				Array.from(view.dom.querySelectorAll('p'), (p) => p.getAttribute('data-list-marker'));
			expect(markers()).toEqual(['', '', '1.', '2.', '3.']);
			view.dispatch(view.state.tr.insertText('Text', must(doc.firstChild).nodeSize + 1));
			expect(markers()).toEqual(['', '1.', '2.', '3.', '4.']);
		} finally {
			view.destroy();
		}
	});

	it('collapses contextual spacing like Word in the editing view (#359)', async () => {
		const fixtures = '../core/docx/layout/fixtures/word-layout-reports';
		const { model } = await loadDocx(
			await readFile(`${fixtures}/issue-359-contextual-spacing-rows.docx`),
		);
		const word = JSON.parse(
			await readFile(`${fixtures}/issue-359-contextual-spacing-rows.json`, 'utf8'),
		) as { gapsPt: Record<string, number> };
		const view = new EditorView(document.createElement('div'), {
			state: EditorState.create({
				doc: modelToDoc(model),
				plugins: [paragraphStylesPlugin(() => model)],
			}),
		});
		try {
			const paragraphs = new Map(
				Array.from(view.dom.querySelectorAll('p'), (p) => [p.textContent ?? '', p]),
			);
			const px = (value: string) => Number.parseFloat(value || '0');
			const gaps = Object.fromEntries(
				Object.keys(word.gapsPt).map((id) => {
					const upper = must(paragraphs.get(`${id}U`));
					const lower = must(paragraphs.get(`${id}L`));
					// Adjacent block margins collapse to the larger one, as Word's auto spacing does.
					const gap = Math.max(px(upper.style.marginBottom), px(lower.style.marginTop));
					return [id, Math.round(gap * 0.75 * 100) / 100];
				}),
			);
			expect(gaps).toEqual(word.gapsPt);
		} finally {
			view.destroy();
		}
	});

	it.each([
		['empty-numbered-paragraphs', 10],
		['control-no-mark-size', 18],
		['control-with-text', 10],
	])('sizes draft markers in %s from the right font (#356)', async (name, size) => {
		const path = `../core/docx/layout/fixtures/word-layout-reports/issue-356-${name}.docx`;
		const { model } = await loadDocx(await readFile(path));
		const view = new EditorView(document.createElement('div'), {
			state: EditorState.create({
				doc: modelToDoc(model),
				plugins: [paragraphStylesPlugin(() => model)],
			}),
		});
		try {
			const markers = Array.from(
				view.dom.querySelectorAll<HTMLParagraphElement>(
					'p[data-list-marker]:not([data-list-marker=""])',
				),
			);
			expect(markers).toHaveLength(12);
			for (const marker of markers) {
				expect(marker.style.getPropertyValue('--dve-marker-size')).toBe(`${(size * 4) / 3}px`);
				expect(marker.style.fontSize).toBe(
					name === 'empty-numbered-paragraphs' ? 'var(--dve-marker-size)' : '',
				);
			}
		} finally {
			view.destroy();
		}
	});

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
		const { model } = await loadDocx(await readFile(tableStyleFixture));
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

	it('renders table font sizes and refreshes them when the compatibility setting changes (#353)', async () => {
		const loaded = await loadDocx(await readFile(tableStyleFixture));
		let { model } = loaded;
		const doc = modelToDoc(model);
		let cellPos = 0;
		doc.descendants((node, pos, parent) => {
			if (cellPos === 0 && node.type.name === 'paragraph' && parent?.type.name === 'tableCell')
				cellPos = pos + 2;
		});
		const view = new EditorView(document.createElement('div'), {
			state: EditorState.create({
				doc,
				selection: TextSelection.create(doc, cellPos),
				plugins: [runStylesPlugin(() => model)],
			}),
		});
		try {
			const readSizes = () =>
				Array.from(view.dom.querySelectorAll('td p'), (p) => {
					const spans = Array.from(
						p.querySelectorAll<HTMLElement>('span[style*="font-size"]'),
					).filter((span) => span.childElementCount === 0);
					return ['Lorem', 'لوريم'].map(
						(text) => spans.find((span) => span.textContent?.includes(text))?.style.fontSize,
					);
				});
			expect.soft(readSizes()).toEqual([
				['16pt', '18pt'],
				['16pt', '18pt'],
				['16pt', '10pt'],
				['9pt', '18pt'],
			]);
			expect.soft(readFontFormat(view.state).size).toBe(16);
			model = { ...model, overrideTableStyleFontSizeAndJustification: true };
			view.updateState(view.state.apply(view.state.tr));
			expect.soft(readSizes()).toEqual([
				['11pt', '12pt'],
				['11pt', '12pt'],
				['11pt', '10pt'],
				['9pt', '12pt'],
			]);
			expect.soft(readFontFormat(view.state).size).toBe(11);
			const edit = must(
				editStyleCatalogs(model, {
					kind: 'paragraph',
					id: 'Normal',
					name: 'Normal',
					paragraphFormatting: {},
					runFormatting: { fontSize: 10, fontSizeComplexScript: 12 },
				}),
			);
			model = { ...model, ...edit.catalogs, overrideTableStyleFontSizeAndJustification: false };
			view.updateState(view.state.apply(view.state.tr));
			expect.soft(readSizes()).toEqual([
				['10pt', '18pt'],
				['10pt', '18pt'],
				['10pt', '10pt'],
				['9pt', '18pt'],
			]);
			expect(readFontFormat(view.state).size).toBe(10);
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
