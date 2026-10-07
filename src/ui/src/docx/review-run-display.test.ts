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
	for (const [atom, selector] of [
		[{ text: '', noteReference: { kind: 'footnote', id: '1' } }, '.dve-note-reference'],
		[{ text: '', break: 'page' }, '.dve-break-marker'],
		[{ text: '', fieldCode: ' PAGE ' }, '.dve-field-marker'],
		[{ text: '\n' }, 'br'],
		[
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
			'[data-docx-image]',
		],
	] as [TextRun, string][])
		it(`projects inline ${selector} formatting without changing identity or history`, () => {
			const instance = editor([
				tracked(
					`<w:rPr xmlns:w="${namespace}"><w:rFonts w:ascii="Georgia"/><w:sz w:val="24"/><w:lang w:val="en-GB"/><w:rtl w:val="0"/><w:i/></w:rPr>`,
					atom,
				),
			]);
			try {
				const source = instance.view.state.doc;
				instance.setMode('original');
				const dom = instance.view.dom.querySelector<HTMLElement>(selector)!;
				expect(dom).not.toBeNull();
				expect(dom.style.fontFamily).toContain('Georgia');
				expect(dom.style.fontSize).toBe('12pt');
				expect(dom.style.fontStyle).toBe('italic');
				expect(dom.style.fontWeight).not.toBe('700');
				expect(dom.style.backgroundColor).toBe('');
				expect(dom.lang).toBe('en-GB');
				expect(dom.dir).toBe('ltr');
				expect(instance.view.state.doc).toBe(source);
				expect(undo(instance.view.state)).toBe(false);
				if (atom.image) {
					expect(dom.getAttribute('width')).toBe('20');
					expect(dom.getAttribute('height')).toBe('30');
				}
				instance.setMode('all');
				const current = instance.view.dom.querySelector<HTMLElement>(selector)!;
				expect(current.style.fontWeight).toBe('700');
				expect(current.style.fontFamily).toContain('Arial');
				expect(current.lang).toBe('fr-FR');
				expect(current.dir).toBe('rtl');
				expect(instance.view.state.doc).toBe(source);
			} finally {
				instance.destroy();
			}
		});
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

	it.each([{ text: 'Text' }, { text: '', noteReference: { kind: 'footnote' as const, id: '1' } }])(
		'reports unavailable history and renders current formatting through the display layer: $text',
		(atom) => {
			const instance = editor([tracked('<rPr xmlns="urn:wrong"/>', atom)]);
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
		},
	);

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
