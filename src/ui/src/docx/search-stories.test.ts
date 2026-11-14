// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { history, undo, undoDepth } from 'prosemirror-history';
import { EditorState, NodeSelection, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import {
	createDocument,
	type Block,
	type DocumentModel,
	type HeaderFooterContent,
	type SectionProperties,
} from 'ooxml-core/docx';
import { schema } from './schema';
import { SearchController, type SearchStoryHost } from './search-controller';
import { modelToDoc } from './model-adapter';
import { storyLabel } from './search-panel-options';

const para = (id: string, text: string, mark?: 'footnote' | 'endnote'): Block => ({
	type: 'paragraph',
	id,
	runs: [...(mark ? [{ text: '1', noteMark: mark }] : []), { text }],
});
const part = (text: string, partName?: string): HeaderFooterContent => ({
	blocks: [para(`h-${text}`, text)],
	...(partName ? { partName } : {}),
});
const section = (extra: Partial<SectionProperties>): SectionProperties =>
	({ ...createDocument().sections?.[0], ...extra }) as SectionProperties;

function makeView(editable = true): EditorView {
	const host = document.createElement('div');
	document.body.append(host);
	const box = (lines: string[], textBoxEditable: boolean) =>
		schema.nodes.image.create({
			unsupported: 'Text box',
			textBoxText: JSON.stringify(lines),
			textBoxEditable,
		});
	const doc = schema.node('doc', null, [
		schema.nodes.paragraph.create({ id: 'body' }, [schema.text('zq body '), box(['zq box'], true)]),
		schema.nodes.paragraph.create({ id: 'locked' }, [box(['zq locked'], false)]),
	]);
	const view: EditorView = new EditorView(host, {
		state: EditorState.create({ doc, plugins: [history()] }),
		editable: () => editable,
		dispatchTransaction(transaction) {
			view.updateState(view.state.apply(transaction));
		},
	});
	return view;
}

function fixture(options: { canEdit?: boolean; editable?: boolean } = {}) {
	const shared = part('zq header', 'word/header1.xml');
	let model: DocumentModel = {
		...createDocument(),
		footnotes: [
			{ id: '1', blocks: [para('f1', ' zq note one', 'footnote')] },
			{ id: '2', blocks: [para('f2', ' zq note two', 'footnote')] },
		],
		endnotes: [{ id: '3', blocks: [para('e1', ' zq end', 'endnote')] }],
		sections: [
			section({ headers: { default: shared }, footers: { first: part('zq footer') } }),
			section({ headers: { default: { ...shared } } }),
		],
	};
	const view = makeView(options.editable ?? true);
	let open: { key: string; view: EditorView } | undefined;
	const shown: string[] = [];
	const host: SearchStoryHost = {
		model: () => model,
		openStory: () => open,
		showStory: vi.fn((story) => shown.push(story.kind)),
		closeStory: vi.fn(() => {
			open = undefined;
		}),
		canEditStories: () => options.canEdit ?? true,
		writeHeaderFooter: vi.fn((story, blocks: Block[]) => {
			const kind = story.kind === 'header' ? 'headers' : 'footers';
			model = {
				...model,
				sections: model.sections!.map((item, index) =>
					index === story.sectionIndex
						? { ...item, [kind]: { ...item[kind], [story.slot]: { blocks } } }
						: item,
				),
			};
		}),
		writeNotes: vi.fn((kind, notes: ReadonlyMap<string, Block[]>) => {
			const key = kind === 'footnote' ? 'footnotes' : 'endnotes';
			model = {
				...model,
				[key]: model[key]!.map((note) => ({ ...note, blocks: notes.get(note.id) ?? note.blocks })),
			};
		}),
	};
	const search = new SearchController(() => view, host);
	return {
		view,
		host,
		search,
		shown,
		model: () => model,
		openStory: (key: string, blocks: Block[]) => {
			const storyHost = document.createElement('div');
			document.body.append(storyHost);
			const storyView: EditorView = new EditorView(storyHost, {
				state: EditorState.create({
					doc: modelToDoc({ ...createDocument(), blocks }),
					plugins: [history()],
				}),
				dispatchTransaction(transaction) {
					storyView.updateState(storyView.state.apply(transaction));
				},
			});
			open = { key, view: storyView };
			return storyView;
		},
	};
}

const text = (blocks: readonly Block[] | undefined) =>
	(blocks ?? [])
		.flatMap((block) => (block.type === 'paragraph' ? block.runs.map((run) => run.text ?? '') : []))
		.join('');

afterEach(() => document.body.replaceChildren());

describe('Find across stories', () => {
	it('walks the body, headers and footers, footnotes, endnotes, then text boxes, and wraps', () => {
		const { search, shown, view } = fixture();
		expect(search.search('zq').count).toBe(8);
		const kinds: string[] = [];
		for (let step = 0; step < 9; step++) {
			search.findNext();
			kinds.push(search.status.story?.kind ?? '');
		}
		expect(kinds).toEqual([
			'main',
			'header',
			'footer',
			'footnote',
			'footnote',
			'endnote',
			'textBoxes',
			'textBoxes',
			'main',
		]);
		// The linked second section's header is the same part: it is searched once.
		expect(shown).toEqual(['header', 'footer', 'footnote', 'footnote', 'endnote']);
		search.findPrevious();
		expect(search.status).toMatchObject({ active: 8, story: { kind: 'textBoxes' } });
		view.destroy();
	});

	it('selects a text box match as the box, and the next Find moves on from it', () => {
		const { search, view } = fixture();
		search.search('zq');
		for (let step = 0; step < 7; step++) search.findNext();
		expect(view.state.selection).toBeInstanceOf(NodeSelection);
		expect(search.status).toMatchObject({ active: 7, story: { kind: 'textBoxes' } });
		search.refresh();
		expect(search.status.active).toBe(7);
		search.findNext();
		expect(search.status).toMatchObject({ active: 8, story: { kind: 'textBoxes' } });
		view.destroy();
	});

	it('never matches or replaces a note number mark', () => {
		const { search, model, view } = fixture();
		expect(search.search('1 zq').count).toBe(0);
		search.search('note one');
		expect(search.replaceAll('first')).toBe(1);
		const blocks = model().footnotes![0]!.blocks;
		expect(blocks[0]?.type === 'paragraph' && blocks[0].runs[0]?.noteMark).toBe('footnote');
		expect(text(blocks)).toBe('1 zq first');
		view.destroy();
	});

	it('replaces every story, one undo step per story, and skips a text box it cannot rewrite', () => {
		const { search, host, model, view } = fixture();
		search.search('zq');
		expect(search.replaceAll('ok')).toBe(7);
		expect(view.state.doc.textContent).toBe('ok body ');
		const boxes: unknown[] = [];
		view.state.doc.descendants((node) => {
			if (node.attrs.textBoxText) boxes.push(JSON.parse(node.attrs.textBoxText));
		});
		expect(boxes).toEqual([['ok box'], ['zq locked']]);
		expect(host.writeNotes).toHaveBeenCalledTimes(2);
		expect(vi.mocked(host.writeNotes).mock.calls[0]?.[1].size).toBe(2);
		expect(text(model().footnotes![1]!.blocks)).toBe('1 ok note two');
		expect(text(model().endnotes![0]!.blocks)).toBe('1 ok end');
		expect(host.writeHeaderFooter).toHaveBeenCalledTimes(2);
		expect(text(model().sections![0]!.headers!.default!.blocks)).toBe('ok header');
		expect(text(model().sections![0]!.footers!.first!.blocks)).toBe('ok footer');
		// The body and the text boxes are separate undo steps (the host writes the other stories).
		expect(undoDepth(view.state)).toBe(2);
		undo(view.state, view.dispatch);
		expect(view.state.doc.textContent).toBe('ok body ');
		expect(view.state.doc.firstChild?.lastChild?.attrs.textBoxText).toBe('["zq box"]');
		undo(view.state, view.dispatch);
		expect(view.state.doc.textContent).toBe('zq body ');
		view.destroy();
	});

	it('finds but does not replace locked stories', () => {
		const { search, host, view } = fixture({ canEdit: false });
		expect(search.search('zq').count).toBe(8);
		expect(search.replaceAll('ok')).toBe(2);
		expect(host.writeNotes).not.toHaveBeenCalled();
		expect(host.writeHeaderFooter).not.toHaveBeenCalled();
		view.destroy();
	});

	it('replaces in an open story editor and resumes Find from its caret', () => {
		const { search, host, openStory, view } = fixture();
		const editor = openStory('endnote:3', [para('e1', ' zq end')]);
		editor.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 2, 4)));
		search.search('zq');
		expect(search.status).toMatchObject({ active: 6, story: { kind: 'endnote' } });
		expect(search.replaceOne('ok')).toBe(true);
		expect(editor.state.doc.textContent).toBe(' ok end');
		expect(host.writeNotes).not.toHaveBeenCalled();
		expect(search.status).toMatchObject({ count: 7, active: 0 });
		search.findNext();
		expect(search.status).toMatchObject({ active: 6, story: { kind: 'textBoxes' } });
		editor.destroy();
		view.destroy();
	});

	it('names the story the way Word does, in every locale', () => {
		expect(storyLabel('en', { kind: 'main' })).toBe('Main Document');
		expect(storyLabel('en', { kind: 'header', slot: 'default', sectionIndex: 0 })).toBe(
			'Header -Section 1-',
		);
		expect(storyLabel('de', { kind: 'footer', slot: 'first', sectionIndex: 1 })).toBe(
			'Fußzeile der ersten Seite -Abschnitt 2-',
		);
		expect(storyLabel('fr', { kind: 'textBoxes' })).toBe('Zones de texte');
		expect(storyLabel('en', undefined)).toBe('');
	});
});
