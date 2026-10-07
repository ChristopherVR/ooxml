// @vitest-environment jsdom
import { expect, it } from 'vitest';
import { createDocument, type TextRun } from 'ooxml-core/docx';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { history, undo } from 'prosemirror-history';
import { modelToDoc, docToModel } from './model-adapter';
import { applyFontFormat, readFontFormat } from './font-format';
import { runStylesPlugin } from './run-styles';
import { clearFormatting } from './ribbon-commands';

const atoms: TextRun[] = [
	{ text: '\n' },
	{ text: '', break: 'page' },
	{ text: '', noteReference: { kind: 'footnote', id: '1' } },
	{ text: '', fieldCode: 'DATE' },
	{
		text: '',
		image: {
			relId: 'rId1',
			partName: 'word/media/a.png',
			contentType: 'image/png',
			widthPx: 10,
			heightPx: 20,
		},
	},
];
for (const atom of atoms)
	it(`edits selected ${atom.text ? 'line break' : Object.keys(atom)[1]} font properties atomically`, () => {
		const model = createDocument();
		model.blocks = [
			{
				type: 'paragraph',
				id: 'p',
				runs: [{ ...atom, fontFamily: 'Georgia', fontSize: 14, smallCaps: true }],
			},
		];
		const doc = modelToDoc(model);
		const view = new EditorView(document.createElement('div'), {
			state: EditorState.create({
				doc,
				selection: TextSelection.create(doc, 1, 2),
				plugins: [history(), runStylesPlugin(() => model)],
			}),
		});
		try {
			expect(readFontFormat(view.state)).toMatchObject({
				family: 'Georgia',
				size: 14,
				smallCaps: true,
			});
			applyFontFormat(view, {
				family: 'Arial',
				size: 18,
				color: '#c00000',
				smallCaps: false,
				spacing: 1.5,
				script: 'subscript',
			});
			expect(docToModel(view.state.doc, model).blocks[0]).toMatchObject({
				runs: [
					{
						...atom,
						fontFamily: 'Arial',
						fontSize: 18,
						color: '#c00000',
						smallCaps: false,
						characterSpacingTwips: 30,
						verticalAlign: 'subscript',
					},
				],
			});
			expect(undo(view.state, view.dispatch)).toBe(true);
			expect(view.state.doc.eq(doc)).toBe(true);
		} finally {
			view.destroy();
		}
	});

it('clears direct formatting while preserving independent comments, hyperlinks and tracked text', () => {
	const model = createDocument();
	model.blocks = [
		{
			type: 'paragraph',
			id: 'p',
			runs: [
				{
					text: 'Text',
					bold: true,
					color: '#c00000',
					smallCaps: true,
					link: { href: 'https://example.com' },
					commentIds: ['c'],
					revision: { kind: 'insert', author: 'Ada', id: 'r' },
				},
			],
		},
	];
	const doc = modelToDoc(model);
	const view = new EditorView(document.createElement('div'), {
		state: EditorState.create({ doc, selection: TextSelection.create(doc, 1, 5) }),
	});
	try {
		clearFormatting(view);
		const cleared = docToModel(view.state.doc, model).blocks[0]!;
		if (cleared.type !== 'paragraph') throw new Error('Expected paragraph');
		expect(cleared.runs[0]).not.toHaveProperty('bold');
		expect(cleared.runs[0]).not.toHaveProperty('color');
		expect(cleared.runs[0]).not.toHaveProperty('smallCaps');
		expect(docToModel(view.state.doc, model).blocks[0]).toMatchObject({
			runs: [
				{
					text: 'Text',
					link: { href: 'https://example.com' },
					commentIds: ['c'],
					revision: { kind: 'insert', author: 'Ada', id: 'r' },
				},
			],
		});
	} finally {
		view.destroy();
	}
});
