// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { history, undo } from 'prosemirror-history';
import type { DocumentModel, ReviewDisplayMode } from 'ooxml-core/docx';
import { schema } from './schema';
import { paragraphStylesPlugin } from './paragraph-styles';
import { runStylesPlugin } from './run-styles';

const namespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const model: DocumentModel = {
	blocks: [],
	warnings: [],
	page: {
		width: 816,
		height: 1056,
		marginTop: 96,
		marginBottom: 96,
		marginLeft: 96,
		marginRight: 96,
	},
	paragraphStyles: {
		docDefaults: {},
		styles: {
			Before: {
				id: 'Before',
				name: 'Before',
				formatting: { align: 'right' },
			},
			After: {
				id: 'After',
				name: 'After',
				formatting: { align: 'center' },
			},
		},
		warnings: [],
	},
	characterStyles: {
		docDefaults: {},
		warnings: [],
		styles: {
			Before: { id: 'Before', type: 'paragraph', formatting: { italic: true } },
			After: { id: 'After', type: 'paragraph', formatting: { bold: true } },
		},
	},
};

function editor(previous: string) {
	let mode: ReviewDisplayMode = 'all';
	const paragraph = schema.node(
		'paragraph',
		{
			id: 'p',
			style: 'After',
			align: 'center',
			indentLeftTwips: 720,
			spacingAfterTwips: 240,
			direction: 'rtl',
			shadingFill: '#FF0000',
			formatRevision: {
				id: 'format',
				kind: 'paragraphChange',
				author: 'Ada',
				previousParagraphPropertiesXml: previous,
			},
		},
		[schema.text('Text')],
	);
	const doc = schema.node('doc', null, [paragraph]);
	const host = document.createElement('div');
	document.body.append(host);
	const view = new EditorView(host, {
		state: EditorState.create({
			doc,
			selection: TextSelection.create(doc, 2),
			plugins: [
				history(),
				runStylesPlugin(
					() => model,
					() => mode,
				),
				paragraphStylesPlugin(
					() => model,
					() => mode,
				),
			],
		}),
	});
	return {
		view,
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

describe('Original paragraph formatting display', () => {
	it('restores prior style inheritance and clears current direct properties without changing history or selection', () => {
		const instance = editor(`<w:pPr xmlns:w="${namespace}"><w:pStyle w:val="Before"/></w:pPr>`);
		try {
			const { view } = instance;
			const source = view.state.doc;
			const selection = view.state.selection.toJSON();
			expect(view.dom.querySelector('p')!.style.textAlign).toBe('center');
			instance.setMode('original');
			const paragraph = view.dom.querySelector('p')!;
			expect(paragraph.style.textAlign).toBe('right');
			expect(paragraph.style.marginLeft).toBe('0px');
			expect(paragraph.style.marginBottom).toBe('0px');
			expect(paragraph.style.backgroundColor).toBe('transparent');
			expect(paragraph.dir).toBe('ltr');
			expect(paragraph.querySelector('span')!.style.fontStyle).toBe('italic');
			expect(paragraph.querySelector('span')!.style.fontWeight).toBe('');
			expect(view.state.doc).toBe(source);
			expect(view.state.selection.toJSON()).toEqual(selection);
			expect(undo(view.state)).toBe(false);
			instance.setMode('final');
			expect(view.dom.querySelector('p')!.style.textAlign).toBe('center');
			expect(view.dom.querySelector('p')!.style.marginLeft).toBe('48px');
			expect(view.state.doc).toBe(source);
		} finally {
			instance.destroy();
		}
	});

	it('keeps the current paragraph visible and exposes a diagnostic for an invalid prior snapshot', () => {
		const instance = editor('<pPr xmlns="urn:wrong"/>');
		try {
			instance.setMode('original');
			const paragraph = instance.view.dom.querySelector('p')!;
			expect(paragraph.style.textAlign).toBe('center');
			expect(paragraph.dataset.reviewFormatError).toBeTruthy();
			instance.setMode('all');
			expect(paragraph.dataset.reviewFormatError).toBeUndefined();
		} finally {
			instance.destroy();
		}
	});
});
