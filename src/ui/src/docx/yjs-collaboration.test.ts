import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { loadDocx } from 'ooxml-core/docx';
import { collectRevisionRanges, rejectRevisionRange } from './review-commands';
import { afterEach, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { createDocument, saveDocx } from 'ooxml-core/docx';
import {
	createCollabSession,
	createMemoryHub,
	transportProvider,
	type CollabSession,
} from 'ooxml-core/collab';
import { toggleTrackChanges, wordYjsPluginKey } from 'ooxml-core/docx/ui';
import type { EditorView } from 'prosemirror-view';
import { DocxEditorElement } from './index';
import './index';
import { editorBindings } from './editor-commands';
import { insertPicture } from './picture-commands';
import { TextSelection } from 'prosemirror-state';
import { addComment, commentIdsAtSelection } from './comment-commands';

const sessions: CollabSession[] = [];
const editors: DocxEditorElement[] = [];
afterEach(() => {
	for (const editor of editors.splice(0)) {
		editor.stopCollaboration(true);
		editor.remove();
	}
	for (const session of sessions.splice(0)) session.destroy();
});
function mount(): DocxEditorElement {
	const editor = document.createElement('docx-editor') as DocxEditorElement;
	document.body.append(editor);
	editors.push(editor);
	return editor;
}
const viewOf = (editor: DocxEditorElement) => (editor as unknown as { view: EditorView }).view;
function pair(viewer = false, filter?: () => boolean) {
	const hub = createMemoryHub(filter ? { filter } : {});
	const join = (name: string, readOnly: boolean) => {
		const session = createCollabSession({
			roomId: 'word',
			provider: transportProvider({ transport: hub.createTransport('word') }),
			user: { name, ...(readOnly ? { role: 'viewer' as const } : {}) },
			heartbeatMs: 0,
			teardown: false,
		});
		sessions.push(session);
		return session;
	};
	return [join('Ada', false), join('Bob', viewer)] as const;
}
function start(
	a: DocxEditorElement,
	b: DocxEditorElement,
	pair: readonly [CollabSession, CollabSession],
) {
	a.startYjsCollaboration(pair[0], { documentId: 'source', initializeIfEmpty: true });
	b.startYjsCollaboration(pair[1], { documentId: 'source' });
}

describe('Word Yjs collaboration', () => {
	it('shares paragraph formatting rejection, peer export and undo/redo', async () => {
		const native = await loadDocx(
			new Uint8Array(
				await readFile(
					resolve('../core/docx/__fixtures__/review-paragraph-formatting/multiple-rejected.docx'),
				),
			),
		);
		const bytes = new Uint8Array(
			await readFile(
				resolve('../core/docx/__fixtures__/review-paragraph-formatting/multiple-tracked.docx'),
			),
		);
		const peers = pair();
		const a = mount();
		const b = mount();
		await a.load(bytes);
		await b.load(bytes);
		start(a, b, peers);
		const first = viewOf(a);
		const before = first.state.doc;
		rejectRevisionRange(first, collectRevisionRanges(before)[0]!);
		for (const editor of [a, b]) {
			const paragraph = editor.documentModel!.blocks[0]!;
			if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
			const { restoredParagraphPropertiesXml: _snapshot, ...actual } = paragraph;
			expect(actual).toEqual(native.model.blocks[0]);
			expect(paragraph.formatRevision).toBeUndefined();
		}
		const reopened = await loadDocx(await b.saveBytes());
		const paragraph = reopened.model.blocks[0]!;
		if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
		expect(paragraph).toEqual(native.model.blocks[0]);
		expect(paragraph.formatRevision).toBeUndefined();
		expect(wordYjsPluginKey.getState(first.state)!.undo()).toBe(true);
		for (const editor of [a, b]) expect(viewOf(editor).state.doc.eq(before)).toBe(true);
		expect(wordYjsPluginKey.getState(first.state)!.redo()).toBe(true);
		expect(collectRevisionRanges(viewOf(b).state.doc)).toEqual([]);
	});
	it.each([true, false])(
		'retains native paragraph formatting history during peer typing with tracking %s',
		async (tracked) => {
			const bytes = new Uint8Array(
				await readFile(
					resolve('../core/docx/__fixtures__/review-paragraph-formatting/multiple-tracked.docx'),
				),
			);
			const loaded = await loadDocx(bytes);
			const paragraph = loaded.model.blocks[0]!;
			if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
			const peers = pair();
			const a = mount();
			const b = mount();
			await a.load(bytes);
			await b.load(bytes);
			start(a, b, peers);
			const first = viewOf(a);
			if (!tracked) toggleTrackChanges(first.state, first.dispatch, first);
			const view = viewOf(b);
			view.dispatch(view.state.tr.insertText('!', 5));
			for (const editor of [a, b]) {
				const actual = editor.documentModel!.blocks[0]!;
				if (actual.type !== 'paragraph') throw new Error('Expected paragraph');
				expect(actual.formatRevision).toEqual(paragraph.formatRevision);
				expect(actual.runs.some((run) => run.revision?.kind === 'insert')).toBe(tracked);
				for (const run of actual.runs)
					expect(run.fontFamilyComplexScript).toBe(paragraph.runs[0]?.fontFamilyComplexScript);
				const reopened = await loadDocx(await editor.saveBytes());
				const exported = reopened.model.blocks[0]!;
				if (exported.type !== 'paragraph') throw new Error('Expected paragraph');
				expect(exported.formatRevision).toEqual(paragraph.formatRevision);
				expect(exported.runs.some((run) => run.revision?.kind === 'insert')).toBe(tracked);
				for (const run of exported.runs)
					expect(run.fontFamilyComplexScript).toBe(paragraph.runs[0]?.fontFamilyComplexScript);
			}
		},
	);
	it('shares imported formatting rejection and its undo without recording remote changes', async () => {
		const bytes = new Uint8Array(
			await readFile(resolve('../core/docx/__fixtures__/review-formatting/multiple-tracked.docx')),
		);
		const peers = pair();
		const a = mount();
		const b = mount();
		await a.load(bytes);
		await b.load(bytes);
		start(a, b, peers);
		const first = viewOf(a);
		const before = first.state.doc;
		const range = collectRevisionRanges(first.state.doc)[0]!;
		expect(range.kind).toBe('formatChange');
		rejectRevisionRange(first, range);
		for (const editor of [a, b]) {
			const paragraph = editor.documentModel!.blocks[0]!;
			if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
			expect(paragraph.runs[0]).toMatchObject({ text: 'Format me', bold: true, color: '#0000FF' });
			expect(paragraph.runs[0]!.italic).toBeUndefined();
			expect(paragraph.runs[0]!.revision).toBeUndefined();
		}
		const reopened = await loadDocx(await b.saveBytes());
		const paragraph = reopened.model.blocks[0]!;
		if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
		expect(paragraph.runs[0]).toMatchObject({ text: 'Format me', bold: true, color: '#0000FF' });
		expect(wordYjsPluginKey.getState(viewOf(a).state)!.undo()).toBe(true);
		for (const editor of [a, b]) expect(viewOf(editor).state.doc.eq(before)).toBe(true);
		expect(wordYjsPluginKey.getState(viewOf(a).state)!.redo()).toBe(true);
		expect(collectRevisionRanges(viewOf(b).state.doc)).toEqual([]);
	});

	it('merges concurrent overlapping anchors and undoes only the local comment mark', () => {
		let deliver = true;
		const peers = pair(false, () => deliver);
		const a = mount();
		const b = mount();
		a.documentModel = {
			...createDocument(),
			blocks: [{ type: 'paragraph', id: 'p', runs: [{ text: 'Shared text' }] }],
		};
		b.documentModel = a.documentModel!;
		start(a, b, peers);
		deliver = false;
		for (const [index, editor] of [a, b].entries()) {
			const view = viewOf(editor);
			view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 7)));
			addComment(view, index ? 'Bob' : 'Ada', 'Note', () => (index ? 'c2' : 'c1'));
		}
		deliver = true;
		a.resyncCollaboration();
		for (const editor of [a, b])
			expect(commentIdsAtSelection(viewOf(editor))).toEqual(['c1', 'c2']);
		expect(viewOf(a).state.doc.toJSON()).toEqual(viewOf(b).state.doc.toJSON());
		const model = a.documentModel!;
		if (model.blocks[0]!.type !== 'paragraph') throw new Error('Expected paragraph');
		expect(model.blocks[0]!.runs[0]!.commentIds).toEqual(['c1', 'c2']);
		expect(editorBindings['Mod-z']!(viewOf(a).state, viewOf(a).dispatch, viewOf(a))).toBe(true);
		for (const editor of [a, b]) expect(commentIdsAtSelection(viewOf(editor))).toEqual(['c2']);
	});
	it('converges after a partition and undoes only the local author', () => {
		let deliver = true;
		const peers = pair(false, () => deliver);
		const a = mount();
		const b = mount();
		start(a, b, peers);
		deliver = false;
		viewOf(a).dispatch(viewOf(a).state.tr.insertText('Ada', 1));
		viewOf(b).dispatch(viewOf(b).state.tr.insertText('Bob', 1));
		deliver = true;
		expect(a.resyncCollaboration()).toBe(true);
		expect(viewOf(a).state.doc.toJSON()).toEqual(viewOf(b).state.doc.toJSON());
		expect(viewOf(a).state.doc.textContent).toContain('Ada');
		expect(viewOf(a).state.doc.textContent).toContain('Bob');
		expect(editorBindings['Mod-z']!(viewOf(a).state, viewOf(a).dispatch, viewOf(a))).toBe(true);
		expect(viewOf(a).state.doc.textContent).toBe('Bob');
		expect(viewOf(b).state.doc.textContent).toBe('Bob');
	});
	it('preserves section settings and local undo across view teardown', () => {
		const peers = pair();
		const a = mount();
		const b = mount();
		start(a, b, peers);
		viewOf(a).dispatch(
			viewOf(a).state.tr.setDocAttribute('pageWidth', 1000).setDocAttribute('marginLeft', 120),
		);
		expect(viewOf(b).state.doc.attrs.pageWidth).toBe(1000);
		expect(viewOf(b).state.doc.attrs.marginLeft).toBe(120);
		wordYjsPluginKey.getState(viewOf(a).state)!.stopCapturing();
		viewOf(a).dispatch(viewOf(a).state.tr.insertText('local', 1));
		a.remove();
		viewOf(b).dispatch(viewOf(b).state.tr.insertText(' remote', 6));
		document.body.append(a);
		expect(viewOf(a).state.doc.textContent).toBe('local remote');
		expect(viewOf(a).state.doc.attrs.pageWidth).toBe(1000);
		expect(editorBindings['Mod-z']!(viewOf(a).state, viewOf(a).dispatch, viewOf(a))).toBe(true);
		expect(viewOf(b).state.doc.textContent).toBe(' remote');
		a.reconnectCollaboration();
		expect(viewOf(a).state.doc.toJSON()).toEqual(viewOf(b).state.doc.toJSON());
	});
	it('rejects local writes for viewer peers and retains opaque DOCX parts on save', async () => {
		const zip = await JSZip.loadAsync(await saveDocx(createDocument()));
		zip.file('custom/preserved.txt', 'package payload');
		const source = await zip.generateAsync({ type: 'uint8array' });
		const a = mount();
		const b = mount();
		await Promise.all([a.load(source), b.load(source)]);
		start(a, b, pair(true));
		b.readOnly = true;
		b.readOnly = false;
		expect(viewOf(b).editable).toBe(false);
		viewOf(b).dispatch(viewOf(b).state.tr.insertText('forbidden', 1));
		expect(viewOf(b).state.doc.textContent).toBe('');
		viewOf(a).dispatch(viewOf(a).state.tr.insertText('مرحبا', 1));
		expect(viewOf(b).state.doc.textContent).toBe('مرحبا');
		const saved = await JSZip.loadAsync(await b.saveBytes());
		expect(await saved.file('custom/preserved.txt')!.async('string')).toBe('package payload');
	});
	it('requires initialized matching rooms and prevents mixing collaboration engines', () => {
		const peers = pair();
		const a = mount();
		const b = mount();
		expect(() => b.startYjsCollaboration(peers[1], { documentId: 'source' })).toThrow(
			/designated creator/,
		);
		a.startYjsCollaboration(peers[0], { documentId: 'source', initializeIfEmpty: true });
		expect(() => b.startYjsCollaboration(peers[1], { documentId: 'different' })).toThrow(
			/matching source/,
		);
		expect(() => a.startCollaboration({ sessionId: 'step', clientId: 'a' })).toThrow(/Stop/);
		expect(() => {
			a.documentModel = createDocument();
		}).toThrow(/Stop/);
		a.stopCollaboration();
		expect(peers[0].status).toBe('connected');
	});
	it('merges concurrent formatting and text', () => {
		let deliver = true;
		const peers = pair(false, () => deliver);
		const a = mount();
		const b = mount();
		start(a, b, peers);
		viewOf(a).dispatch(viewOf(a).state.tr.insertText('text', 1));
		wordYjsPluginKey.getState(viewOf(a).state)!.stopCapturing();
		deliver = false;
		viewOf(a).dispatch(
			viewOf(a).state.tr.addMark(1, 5, viewOf(a).state.schema.marks.bold!.create()),
		);
		viewOf(b).dispatch(
			viewOf(b)
				.state.tr.addMark(1, 5, viewOf(b).state.schema.marks.italic!.create())
				.insertText('!', 5),
		);
		deliver = true;
		a.resyncCollaboration();
		expect(viewOf(a).state.doc.toJSON()).toEqual(viewOf(b).state.doc.toJSON());
		expect(viewOf(a).state.doc.textContent).toBe('text!');
		expect(viewOf(a).state.doc.firstChild!.firstChild!.marks.map((mark) => mark.type.name)).toEqual(
			['bold', 'italic'],
		);
		expect(editorBindings['Mod-z']!(viewOf(a).state, viewOf(a).dispatch, viewOf(a))).toBe(true);
		expect(viewOf(b).state.doc.firstChild!.firstChild!.marks.map((mark) => mark.type.name)).toEqual(
			['italic'],
		);
		expect(viewOf(b).state.doc.textContent).toBe('text!');
	});
	it('receives tracked insertions without assigning them to the receiving author', () => {
		const peers = pair();
		const a = mount();
		const b = mount();
		const model = createDocument();
		model.trackChanges = true;
		a.documentModel = structuredClone(model);
		b.documentModel = structuredClone(model);
		a.reviewAuthor = 'Ada';
		b.reviewAuthor = 'Bob';
		start(a, b, peers);
		viewOf(a).dispatch(viewOf(a).state.tr.insertText('tracked', 1));
		expect(viewOf(a).state.doc.toJSON()).toEqual(viewOf(b).state.doc.toJSON());
		const marks = viewOf(b).state.doc.firstChild!.firstChild!.marks;
		expect(marks.filter((mark) => mark.type.name === 'insertion')).toHaveLength(1);
		expect(marks.find((mark) => mark.type.name === 'insertion')!.attrs.author).toBe('Ada');
	});
	it('undoes page settings while excluding non-history updates and command probes', () => {
		const peers = pair();
		const a = mount();
		const b = mount();
		start(a, b, peers);
		const width = viewOf(a).state.doc.attrs.pageWidth;
		viewOf(a).dispatch(viewOf(a).state.tr.setDocAttribute('pageWidth', 1000));
		expect(editorBindings['Mod-z']!(viewOf(a).state)).toBe(true);
		expect(viewOf(a).state.doc.attrs.pageWidth).toBe(1000);
		expect(editorBindings['Mod-z']!(viewOf(a).state, viewOf(a).dispatch, viewOf(a))).toBe(true);
		expect(viewOf(b).state.doc.attrs.pageWidth).toBe(width);
		viewOf(a).dispatch(
			viewOf(a).state.tr.setDocAttribute('pageColor', 'FFFF00').setMeta('addToHistory', false),
		);
		expect(editorBindings['Mod-z']!(viewOf(a).state)).toBe(false);
		expect(viewOf(b).state.doc.attrs.pageColor).toBe('FFFF00');
		a.readOnly = true;
		viewOf(a).dispatch(viewOf(a).state.tr.insertText('forbidden', 1));
		expect(viewOf(b).state.doc.textContent).toBe('');
	});
	it('exports both coauthors’ picture bytes and retains them after stopping while detached', async () => {
		let deliver = true;
		const peers = pair(false, () => deliver);
		const a = mount();
		const b = mount();
		start(a, b, peers);
		deliver = false;
		const names: string[] = [];
		for (const [index, editor] of [a, b].entries()) {
			const media = wordYjsPluginKey.getState(viewOf(editor).state)!.media;
			const name = media.partName('image/png');
			names.push(name);
			media.publish(name, {
				contentType: 'image/png',
				bytes: new Uint8Array([137, 80, 78, index]),
			});
			insertPicture(viewOf(editor), {
				relId: '',
				partName: name,
				contentType: 'image/png',
				widthPx: 10,
				heightPx: 10,
			});
		}
		deliver = true;
		a.resyncCollaboration();
		expect(viewOf(a).state.doc.toJSON()).toEqual(viewOf(b).state.doc.toJSON());
		b.remove();
		viewOf(a).dispatch(viewOf(a).state.tr.insertText('after detach', 1));
		const liveZip = await JSZip.loadAsync(await b.saveBytes());
		expect(await liveZip.file('word/document.xml')!.async('string')).toContain('after detach');
		b.stopCollaboration();
		const zip = await JSZip.loadAsync(await b.saveBytes());
		for (const [index, name] of names.entries())
			expect(await zip.file(name)!.async('uint8array')).toEqual(
				new Uint8Array([137, 80, 78, index]),
			);
		expect(await zip.file('word/document.xml')!.async('string')).toContain('after detach');
	});
	it('refreshes a mounted picture when its media arrives after the document reference', () => {
		const create = Object.getOwnPropertyDescriptor(URL, 'createObjectURL');
		const revoke = Object.getOwnPropertyDescriptor(URL, 'revokeObjectURL');
		Object.defineProperty(URL, 'createObjectURL', {
			configurable: true,
			value: () => 'blob:arrived',
		});
		Object.defineProperty(URL, 'revokeObjectURL', { configurable: true, value: () => {} });
		try {
			const peers = pair();
			const a = mount();
			const b = mount();
			start(a, b, peers);
			const media = wordYjsPluginKey.getState(viewOf(a).state)!.media;
			const name = media.partName('image/png');
			insertPicture(viewOf(a), {
				relId: '',
				partName: name,
				contentType: 'image/png',
				widthPx: 10,
				heightPx: 10,
			});
			const image = b.shadowRoot!.querySelector<HTMLImageElement>('img[data-docx-image]')!;
			expect(image.classList.contains('dve-image-missing')).toBe(true);
			media.publish(name, { contentType: 'image/png', bytes: new Uint8Array([137, 80, 78, 71]) });
			expect(image.getAttribute('src')).toBe('blob:arrived');
			expect(image.classList.contains('dve-image-missing')).toBe(false);
		} finally {
			if (create) Object.defineProperty(URL, 'createObjectURL', create);
			else Reflect.deleteProperty(URL, 'createObjectURL');
			if (revoke) Object.defineProperty(URL, 'revokeObjectURL', revoke);
			else Reflect.deleteProperty(URL, 'revokeObjectURL');
		}
	});
});
