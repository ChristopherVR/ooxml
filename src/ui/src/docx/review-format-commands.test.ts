import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { createDocument, loadDocx, type Paragraph } from 'ooxml-core/docx';
import { EditorState } from 'prosemirror-state';
import { DecorationSet, EditorView } from 'prosemirror-view';
import { history, undo, redo } from 'prosemirror-history';
import {
	collectRevisionRanges,
	acceptRevisionRange,
	rejectRevisionRange,
	goToNextChange,
} from 'ooxml-core/docx/ui';
import { modelToDoc, docToModel } from './model-adapter';
import { trackChangesPlugin } from './track-changes-mode';
import { reviewDisplayPlugin, type ReviewDisplayMode } from './review-display';

const views: EditorView[] = [];
afterEach(() => {
	for (const view of views.splice(0)) view.destroy();
	document.body.replaceChildren();
});
for (const name of ['bold', 'multiple'])
	describe(`review imported ${name} formatting`, () => {
		for (const mode of ['accept', 'reject'] as const)
			it(`${mode} preserves text and supports undo/redo with recording enabled`, async () => {
				const bytes = await readFile(
					resolve('../core/docx/__fixtures__/review-formatting', `${name}-tracked.docx`),
				);
				const loaded = await loadDocx(new Uint8Array(bytes));
				const host = document.body.appendChild(document.createElement('div'));
				const view = new EditorView(host, {
					state: EditorState.create({
						doc: modelToDoc(loaded.model),
						plugins: [
							history(),
							trackChangesPlugin(
								() => 'Reviewer',
								() => true,
							),
						],
					}),
				});
				views.push(view);
				const before = view.state.doc;
				const ranges = collectRevisionRanges(before);
				expect(ranges).toHaveLength(1);
				expect(ranges[0]!.kind).toBe('formatChange');
				expect(goToNextChange(view)).toBe(true);
				expect(view.state.selection.from).toBe(ranges[0]!.from);
				(mode === 'accept' ? acceptRevisionRange : rejectRevisionRange)(view, ranges[0]!);
				expect(view.state.doc.textContent).toBe('Format me');
				expect(collectRevisionRanges(view.state.doc)).toEqual([]);
				const model = docToModel(view.state.doc, loaded.model);
				const actual = model.blocks[0] as Paragraph;
				if (mode === 'reject') {
					const native = await loadDocx(
						new Uint8Array(
							await readFile(
								resolve('../core/docx/__fixtures__/review-formatting', `${name}-rejected.docx`),
							),
						),
					);
					expect(actual.runs.map(({ restoredRunPropertiesXml: _xml, ...run }) => run)).toEqual(
						(native.model.blocks[0] as Paragraph).runs,
					);
				} else
					expect(actual.runs).toEqual(
						(loaded.model.blocks[0] as Paragraph).runs.map(
							({ revision: _revision, ...run }) => run,
						),
					);
				const reopened = await loadDocx(await loaded.save(model));
				expect((reopened.model.blocks[0] as Paragraph).runs.some((run) => run.revision)).toBe(
					false,
				);
				expect(undo(view.state, view.dispatch)).toBe(true);
				expect(view.state.doc.eq(before)).toBe(true);
				expect(redo(view.state, view.dispatch)).toBe(true);
				expect(collectRevisionRanges(view.state.doc)).toEqual([]);
			});
	});

describe('format review command boundaries', () => {
	it('shows the formatting indicator only in All markup', () => {
		const model = createDocument();
		model.blocks = [
			{
				type: 'paragraph',
				id: 'p',
				runs: [
					{
						text: 'Changed',
						bold: true,
						revision: {
							kind: 'formatChange',
							id: 'f',
							author: 'Ada',
							previousRunPropertiesXml:
								'<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
						},
					},
				],
			},
		];
		let mode: ReviewDisplayMode = 'all';
		const plugin = reviewDisplayPlugin(() => mode);
		const state = EditorState.create({ doc: modelToDoc(model), plugins: [plugin] });
		const decorations = plugin.props.decorations!.call(plugin, state);
		expect(decorations).toBeInstanceOf(DecorationSet);
		expect((decorations as DecorationSet).find()).toHaveLength(1);
		for (const next of ['final', 'simple', 'original'] as const) {
			mode = next;
			expect(plugin.props.decorations!.call(plugin, state)).toBeNull();
		}
	});
	it('restores formatting while retaining links and overlapping comment anchors', () => {
		const model = createDocument();
		model.blocks = [
			{
				type: 'paragraph',
				id: 'p',
				runs: [
					{
						text: 'Linked',
						bold: true,
						link: { href: 'https://example.com' },
						commentIds: ['first', 'second'],
						revision: {
							kind: 'formatChange',
							id: 'f',
							author: 'Ada',
							previousRunPropertiesXml:
								'<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:i/></w:rPr>',
						},
					},
				],
			},
		];
		const host = document.body.appendChild(document.createElement('div'));
		const view = new EditorView(host, { state: EditorState.create({ doc: modelToDoc(model) }) });
		views.push(view);
		rejectRevisionRange(view, collectRevisionRanges(view.state.doc)[0]!);
		const run = (docToModel(view.state.doc, model).blocks[0] as Paragraph).runs[0]!;
		expect(run).toMatchObject({
			text: 'Linked',
			italic: true,
			link: { href: 'https://example.com' },
			commentIds: ['first', 'second'],
		});
		expect(run.bold).toBeUndefined();
		expect(run.revision).toBeUndefined();
	});
	it('blocks formatting resolution in a read-only view', () => {
		const model = createDocument();
		model.blocks = [
			{
				type: 'paragraph',
				id: 'p',
				runs: [
					{
						text: 'Read only',
						bold: true,
						revision: {
							kind: 'formatChange',
							id: 'f',
							author: 'Ada',
							previousRunPropertiesXml:
								'<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
						},
					},
				],
			},
		];
		const host = document.body.appendChild(document.createElement('div'));
		const view = new EditorView(host, {
			state: EditorState.create({ doc: modelToDoc(model) }),
			editable: () => false,
		});
		views.push(view);
		const before = view.state.doc;
		const range = collectRevisionRanges(before)[0]!;
		acceptRevisionRange(view, range);
		rejectRevisionRange(view, range);
		expect(view.state.doc).toBe(before);
	});
});
