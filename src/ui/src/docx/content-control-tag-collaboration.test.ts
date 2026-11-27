// @vitest-environment jsdom
import { readFile } from 'node:fs/promises';
import { EditorState, TextSelection } from 'prosemirror-state';
import { ReplaceStep } from 'prosemirror-transform';
import { EditorView } from 'prosemirror-view';
import { afterEach, describe, expect, it, vi } from 'vitest';
import {
	createCollabSession,
	createMemoryHub,
	transportProvider,
	type CollabSession,
} from 'ooxml-core/collab';
import { loadDocx } from 'ooxml-core/docx';
import {
	changesContentControlTags,
	contentControlPlugin,
	controlSpans,
	isWordYjsRemoteTransaction,
	updateControlProperties,
	WordYjsCollaboration,
	wordYjsUndo,
} from 'ooxml-core/docx/ui';
import type { TagEdge } from './content-control-tag-widgets';
import { contentControlViewPlugin } from './content-control-view';
import { docToModel, modelToDoc } from './model-adapter';

const cleanups: (() => void)[] = [];
afterEach(() => {
	for (const cleanup of cleanups.splice(0).reverse()) cleanup();
});

async function peers(suffix = '') {
	const loaded = await loadDocx(
		await readFile(
			new URL(
				'../../../core/docx/layout/fixtures/word-layout-reports/issue-349-content-control-tags' +
					suffix +
					'.docx',
				import.meta.url,
			),
		),
	);
	const initial = modelToDoc(loaded.model);
	const hub = createMemoryHub();
	const connect = (name: string) => {
		const session = createCollabSession({
			roomId: 'word',
			provider: transportProvider({ transport: hub.createTransport('word') }),
			user: { name },
			heartbeatMs: 0,
			teardown: false,
		});
		cleanups.push(() => session.destroy());
		return session;
	};
	const a = connect('A');
	const b = connect('B');
	const mount = (session: CollabSession, initializeIfEmpty: boolean) => {
		const binding = new WordYjsCollaboration(session, initial, {
			documentId: 'source',
			initializeIfEmpty,
		});
		cleanups.push(() => binding.destroy());
		const { doc, plugins } = binding.state(initial.type.schema);
		const tags = contentControlViewPlugin();
		const host = document.body.appendChild(document.createElement('div'));
		cleanups.push(() => host.remove());
		const view = new EditorView(host, {
			state: EditorState.create({ doc, plugins: [contentControlPlugin(), tags, ...plugins] }),
		});
		cleanups.push(() => view.destroy());
		return { view, tags };
	};
	return { a: mount(a, true), b: mount(b, false), loaded };
}

function firstControl(view: EditorView) {
	const span = controlSpans(view.state.doc).find(({ level }) => level === 'inline');
	if (!span) throw new Error('Missing inline control');
	return span;
}

function firstControlText(view: EditorView) {
	const { from, to } = firstControl(view);
	return view.state.doc.textBetween(from, to);
}

function selectInsideTag(view: EditorView, edge: TagEdge = 'end') {
	const { from, to } = firstControl(view);
	for (const pos of edge === 'start' ? [from + 1, from] : [to - 1, to])
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, pos)));
}

describe('collaborative content control tag caret', () => {
	it.each([
		['start', 'ipsum ', 'ipsum alpha'],
		['end', ' ipsum', 'alpha ipsum'],
	] as const)(
		'keeps typing inside the %s tag after a peer edits preceding text',
		async (edge, text, expected) => {
			const { a, b, loaded } = await peers();
			selectInsideTag(a.view, edge);
			b.view.dispatch(b.view.state.tr.insertText('Lorem ', 2));
			expect(a.view.state.doc.textContent).toBe(b.view.state.doc.textContent);
			const { from, to } = firstControl(a.view);
			expect(a.view.state.selection.from).toBe(edge === 'start' ? from : to);
			a.view.dispatch(a.view.state.tr.insertText(text));
			expect(firstControlText(a.view)).toBe(expected);
			expect(firstControlText(b.view)).toBe(expected);

			const saved = await loaded.save(docToModel(b.view.state.doc, loaded.model));
			const reopened = modelToDoc((await loadDocx(saved)).model);
			const span = controlSpans(reopened)[0];
			expect(span?.control.appearance).toBe('tags');
			expect(span && reopened.textBetween(span.from, span.to)).toBe(expected);
		},
	);

	it('keeps typing inside the end tag after a peer edits the same control', async () => {
		const { a, b } = await peers();
		selectInsideTag(a.view);
		b.view.dispatch(b.view.state.tr.insertText('Lorem', firstControl(b.view).from + 1));
		expect(firstControlText(a.view)).toBe('aLoremlpha');
		expect(a.view.state.selection.from).toBe(firstControl(a.view).to);
		a.view.dispatch(a.view.state.tr.insertText(' ipsum'));
		expect(firstControlText(a.view)).toBe('aLoremlpha ipsum');
		expect(firstControlText(b.view)).toBe('aLoremlpha ipsum');
	});

	it('uses the control id when a peer changes the tag title', async () => {
		const { a, b } = await peers();
		selectInsideTag(a.view);
		const span = firstControl(b.view);
		b.view.dispatch(
			updateControlProperties(b.view.state, span, { ...span.control, alias: 'Lorem' }),
		);
		expect(a.view.dom.querySelector('[data-tag-title=Lorem]')).not.toBeNull();
		a.view.dispatch(a.view.state.tr.insertText(' ipsum'));
		expect(firstControlText(b.view)).toBe('alpha ipsum');
	});

	it('clears the inside side when the restored edge belongs to another control', async () => {
		const { a, b } = await peers();
		selectInsideTag(a.view);
		const span = firstControl(b.view);
		b.view.dispatch(updateControlProperties(b.view.state, span, { ...span.control, id: '999' }));
		expect(firstControl(a.view).control.id).toBe('999');
		expect(a.view.state.selection.from).toBe(firstControl(a.view).to);
		expect(a.tags.getState(a.view.state)?.inside).toBeUndefined();
		a.view.dispatch(a.view.state.tr.insertText(' ipsum'));
		expect(firstControlText(b.view)).toBe('alpha');
	});

	it('clears the inside side and removes the widgets when a peer deletes the control', async () => {
		const { a, b } = await peers();
		selectInsideTag(a.view);
		const { control, from, to } = firstControl(b.view);
		b.view.dispatch(b.view.state.tr.delete(from, to));
		expect(a.tags.getState(a.view.state)?.inside).toBeUndefined();
		expect(a.view.dom.querySelectorAll('.dve-sdt-tag')).toHaveLength(238);
		expect(a.view.state.selection.from).toBe(from);
		a.view.dispatch(a.view.state.tr.insertText('Lorem'));
		expect(a.view.state.doc.textBetween(from, from + 5)).toBe('Lorem');
		expect(controlSpans(a.view.state.doc).some((span) => span.control.id === control.id)).toBe(
			false,
		);
	});

	it('undoes only a local tag rename after a peer edit and keeps subsequent typing inside', async () => {
		const { a, b } = await peers();
		selectInsideTag(a.view);
		const span = firstControl(a.view);
		a.view.dispatch(
			updateControlProperties(a.view.state, span, { ...span.control, alias: 'Ipsum' }),
		);
		expect(firstControl(b.view).control.alias).toBe('Ipsum');
		b.view.dispatch(b.view.state.tr.insertText('Lorem', 2));
		expect(wordYjsUndo(a.view.state, a.view.dispatch, a.view)).toBe(true);
		for (const { view } of [a, b]) {
			expect(firstControl(view).control.alias).toBe(span.control.alias);
			expect(view.state.doc.textBetween(2, 7)).toBe('Lorem');
		}
		expect(a.view.state.selection.from).toBe(firstControl(a.view).to);
		a.view.dispatch(a.view.state.tr.insertText(' dolor'));
		expect(firstControlText(b.view)).toBe('alpha dolor');
	});
});

it('skips the tag rebuild in the untagged fixture until a peer adds tags', async () => {
	const { a, b } = await peers('-control');
	const before = a.tags.getState(a.view.state);
	b.view.dispatch(b.view.state.tr.insertText('Lorem ', 2));
	expect(a.view.state.doc.textContent).toBe(b.view.state.doc.textContent);
	expect(a.tags.getState(a.view.state)).toBe(before);
	expect(a.view.dom.querySelectorAll('.dve-sdt-tag')).toHaveLength(0);
	const span = firstControl(b.view);
	b.view.dispatch(
		updateControlProperties(b.view.state, span, { ...span.control, appearance: 'tags' }),
	);
	expect(a.view.dom.querySelectorAll('.dve-sdt-tag')).toHaveLength(2);
});

it('recognizes that a sync replacement invalidates existing tags without scanning its slice', async () => {
	const { a, b } = await peers();
	const dispatch = vi.spyOn(a.view, 'dispatch');
	let sync;
	try {
		b.view.dispatch(b.view.state.tr.insertText('Lorem ', 2));
		sync = dispatch.mock.calls.map(([tr]) => tr).find(isWordYjsRemoteTransaction);
	} finally {
		dispatch.mockRestore();
	}
	const step = sync?.steps[0];
	if (!sync || !(step instanceof ReplaceStep)) throw new Error('Missing sync replacement');
	const scan = vi.spyOn(step.slice.content, 'descendants');
	try {
		expect(changesContentControlTags(sync, true)).toBe(true);
		expect(scan).not.toHaveBeenCalled();
	} finally {
		scan.mockRestore();
	}
});
