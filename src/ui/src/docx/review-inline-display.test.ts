// @vitest-environment jsdom
import { expect, it } from 'vitest';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { createDocument, loadDocx, saveDocx, type TextRun } from 'ooxml-core/docx';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { modelToDoc, docToModel } from './model-adapter';
import { reviewDisplayPlugin, type ReviewDisplayMode } from './review-display';
import { imageNodeView, ImageMediaCache } from './image-media';
import { acceptAllChanges, rejectAllChanges } from './review-commands';

const atoms: [string, TextRun][] = [
	['.dve-break-marker', { text: '', break: 'page' }],
	[
		'.dve-picture',
		{
			text: '',
			image: {
				relId: 'rId1',
				partName: 'word/media/image.png',
				contentType: 'image/png',
				widthPx: 20,
				heightPx: 20,
			},
		},
	],
	['.dve-note-reference', { text: '', noteReference: { kind: 'footnote', id: '1' } }],
	[
		'.dve-equation',
		{
			text: '',
			equation: {
				omml: '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x</m:t></m:r></m:oMath>',
				display: false,
			},
		},
	],
	['.dve-field-marker', { text: '', fieldChar: 'begin' }],
];
for (const [selector, atom] of atoms)
	it(`filters ${selector} revisions without losing source properties or positions`, () => {
		const model = createDocument();
		model.blocks = [
			{
				type: 'paragraph',
				id: 'p',
				runs: [
					{
						...atom,
						revision: { id: 'insert', kind: 'moveTo', author: 'Ada', move: { name: 'move' } },
					},
					{ text: 'Keep' },
					{
						...atom,
						revision: { id: 'delete', kind: 'moveFrom', author: 'Ada', move: { name: 'move' } },
					},
				],
			},
		];
		let mode: ReviewDisplayMode = 'all';
		const host = document.createElement('div');
		document.body.append(host);
		const view = new EditorView(host, {
			state: EditorState.create({
				doc: modelToDoc(model),
				plugins: [reviewDisplayPlugin(() => mode)],
			}),
			nodeViews: { image: imageNodeView(new ImageMediaCache(() => undefined)) },
		});
		try {
			const source = view.state.doc;
			view.dispatch(view.state.tr.setSelection(TextSelection.create(source, 2)));
			const selection = view.state.selection;
			for (const value of ['original', 'final', 'simple', 'all'] as const) {
				mode = value;
				view.dispatch(view.state.tr.setMeta('addToHistory', false));
				const elements = view.dom.querySelectorAll(selector);
				expect(elements).toHaveLength(2);
				expect(elements[0]!.classList.contains('dve-revision-hidden')).toBe(value === 'original');
				expect(elements[1]!.classList.contains('dve-revision-hidden')).toBe(
					value === 'final' || value === 'simple',
				);
				expect(view.state.doc).toBe(source);
				expect(view.state.selection.eq(selection)).toBe(true);
				expect(docToModel(view.state.doc, model).blocks).toEqual(model.blocks);
			}
		} finally {
			view.destroy();
			host.remove();
		}
	});

it('retains page-break revision history through an editor text edit and package round trip', async () => {
	const model = createDocument();
	model.blocks = [
		{
			type: 'paragraph',
			id: 'p',
			runs: [
				{ text: 'Before' },
				{ text: '', break: 'page', revision: { id: 'r', kind: 'insert', author: 'Ada' } },
			],
		},
	];
	const doc = modelToDoc(model);
	const state = EditorState.create({ doc });
	const next = docToModel(state.tr.insertText('! ', 1).doc, model);
	const reopened = (await loadDocx(await saveDocx(next))).model;
	expect(reopened.blocks[0]).toMatchObject({
		runs: [{ text: '! Before' }, { break: 'page', revision: { kind: 'insert', author: 'Ada' } }],
	});
});

for (const name of [
	'picture-insert',
	'picture-delete',
	'note-insert',
	'note-delete',
	'break-delete',
])
	for (const mode of ['accept', 'reject'] as const)
		it(`exports native ${name} after editor ${mode} without leaking removed picture properties`, async () => {
			const loaded = await loadDocx(
				await readFile(resolve('../core/docx/__fixtures__/review-inline', `${name}-tracked.docx`)),
			);
			const host = document.body.appendChild(document.createElement('div'));
			const view = new EditorView(host, {
				state: EditorState.create({ doc: modelToDoc(loaded.model) }),
			});
			try {
				expect((mode === 'accept' ? acceptAllChanges : rejectAllChanges)(view)).toBe(true);
				const model = docToModel(view.state.doc, loaded.model);
				const bytes = await loaded.save(model);
				const reopened = (await loadDocx(bytes)).model;
				const paragraphs = reopened.blocks.flatMap((block) =>
					block.type === 'paragraph' ? [block] : [],
				);
				expect(paragraphs.flatMap((paragraph) => paragraph.runs).some((run) => run.revision)).toBe(
					false,
				);
				const zip = await JSZip.loadAsync(bytes);
				if (name.startsWith('picture')) {
					const retained = (name === 'picture-insert') === (mode === 'accept');
					const xml = await zip.file('word/document.xml')!.async('string');
					if (retained) {
						expect(xml).toContain('noProof');
						expect(await zip.file('word/media/image1.png')!.async('uint8array')).toEqual(
							loaded.media!.get('word/media/image1.png'),
						);
					} else expect(xml).not.toContain('noProof');
					expect(
						paragraphs
							.flatMap((paragraph) => paragraph.runs)
							.filter((run) => !run.image)
							.every((run) => !run.sourceRunPropertiesXml?.includes('noProof')),
					).toBe(true);
				}
			} finally {
				view.destroy();
				host.remove();
			}
		});
