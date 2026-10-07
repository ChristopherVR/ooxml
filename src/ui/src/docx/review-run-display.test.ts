// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { history, undo } from 'prosemirror-history';
import { createDocument, type TextRun, type ReviewDisplayMode } from 'ooxml-core/docx';
import { modelToDoc, docToModel } from './model-adapter';
import { runStylesPlugin } from './run-styles';
import { reviewRunMarksPlugin } from './review-run-marks';

const namespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const tracked = (prior: string, extra: Partial<TextRun> = {}): TextRun => ({
	text: 'Text',
	bold: true,
	italic: true,
	underline: true,
	strike: true,
	highlight: 'yellow',
	verticalAlign: 'superscript',
	fontSize: 18,
	fontFamily: 'Arial',
	color: '#FF0000',
	language: 'fr-FR',
	rtl: true,
	formatRevision: { id: 'f', kind: 'formatChange', author: 'Ada', previousRunPropertiesXml: prior },
	...extra,
});
function editor(runs: TextRun[]) {
	let mode: ReviewDisplayMode = 'all';
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p', runs }];
	const host = document.createElement('div');
	document.body.append(host);
	const view = new EditorView(host, {
		state: EditorState.create({
			doc: modelToDoc(model),
			plugins: [
				history(),
				reviewRunMarksPlugin(() => mode),
				runStylesPlugin(
					() => model,
					() => mode,
				),
			],
		}),
	});
	return {
		view,
		model,
		setMode(value: ReviewDisplayMode) {
			mode = value;
			view.dispatch(view.state.tr.setMeta('addToHistory', false));
		},
		destroy() {
			view.destroy();
			host.remove();
		},
	};
}

describe('Original run formatting rendering', () => {
	it('neutralizes current outer marks and renders complete prior properties without changing source or undo', () => {
		const instance = editor([
			tracked(
				`<w:rPr xmlns:w="${namespace}"><w:rFonts w:ascii="Georgia"/><w:i/><w:sz w:val="24"/><w:color w:val="0000FF"/><w:lang w:val="en-GB"/><w:rtl w:val="0"/></w:rPr>`,
			),
		]);
		try {
			const source = instance.view.state.doc;
			instance.setMode('original');
			const dom = instance.view.dom;
			expect(dom.querySelector('u')!.style.textDecoration).toBe('none');
			expect(dom.querySelector('sup')!.style.verticalAlign).toBe('baseline');
			const leaf = dom.querySelector<HTMLElement>('span[lang="en-GB"]')!;
			expect(leaf.dir).toBe('ltr');
			expect(leaf.style.fontFamily).toContain('Georgia');
			expect(leaf.style.fontSize).toBe('12pt');
			expect(leaf.style.fontStyle).toBe('italic');
			expect(leaf.style.color).toBe('rgb(0, 0, 255)');
			expect(leaf.style.textDecorationLine).toBe('');
			expect(instance.view.state.doc).toBe(source);
			expect(undo(instance.view.state)).toBe(false);
			instance.setMode('all');
			expect(dom.querySelector('u')!.getAttribute('style')).toBeNull();
			expect(dom.querySelector('sup')!.getAttribute('style')).toBeNull();
			expect(dom.querySelector('span[lang="fr-FR"]')).not.toBeNull();
			expect(instance.view.state.doc).toBe(source);
		} finally {
			instance.destroy();
		}
	});

	it('keeps adjacent untracked bold rendering when a shared outer bold mark is neutralized', () => {
		const instance = editor([
			tracked(`<w:rPr xmlns:w="${namespace}"/>`, {
				italic: false,
				underline: false,
				strike: false,
			}),
			{ text: 'Bold', bold: true },
		]);
		try {
			instance.setMode('original');
			const spans = [...instance.view.dom.querySelectorAll<HTMLElement>('span[style]')];
			expect(
				spans.some((span) => span.textContent === 'Bold' && span.style.fontWeight === '700'),
			).toBe(true);
			expect(instance.view.dom.querySelector('strong')!.style.fontWeight).toBe('inherit');
		} finally {
			instance.destroy();
		}
	});

	it('reports unavailable history and renders current formatting through the display layer', () => {
		const instance = editor([tracked('<rPr xmlns="urn:wrong"/>')]);
		try {
			instance.setMode('original');
			const leaf = instance.view.dom.querySelector<HTMLElement>('[data-review-format-error]')!;
			expect(leaf).not.toBeNull();
			expect(leaf.style.fontWeight).toBe('700');
			expect(leaf.style.textDecorationLine).toContain('underline');
			expect(leaf.style.backgroundColor).toBe('rgb(255, 255, 0)');
		} finally {
			instance.destroy();
		}
	});

	it('observes real text edits without importing neutral display styles into the document', async () => {
		const instance = editor([tracked(`<w:rPr xmlns:w="${namespace}"/>`)]);
		try {
			instance.setMode('original');
			const walker = document.createTreeWalker(instance.view.dom, NodeFilter.SHOW_TEXT);
			const text = walker.nextNode()!;
			text.nodeValue = 'Edited';
			await new Promise((resolve) => setTimeout(resolve, 30));
			const paragraph = docToModel(instance.view.state.doc, instance.model).blocks[0]!;
			if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
			expect(paragraph.runs[0]).toMatchObject({
				text: 'Edited',
				bold: true,
				italic: true,
				underline: true,
				highlight: 'yellow',
				fontSize: 18,
			});
			expect(paragraph.runs[0]!.formatRevision).toEqual(
				instance.model.blocks[0]!.type === 'paragraph'
					? instance.model.blocks[0]!.runs[0]!.formatRevision
					: undefined,
			);
		} finally {
			instance.destroy();
		}
	});
});
