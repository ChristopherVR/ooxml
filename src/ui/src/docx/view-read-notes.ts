import type { DocumentModel, Note } from 'ooxml-core/docx';
import type { Node } from 'prosemirror-model';
import { Plugin, PluginKey } from 'prosemirror-state';
import { Decoration, DecorationSet, type EditorView } from 'prosemirror-view';
import { translate, translateTemplate, type EditorLocale } from './localization';
import { renderStoryPreview } from './story-render';

export interface ReadNoteReference {
	pos: number;
	kind: 'footnote' | 'endnote';
	id: string;
	label: string;
}

/** Every footnote and endnote reference in the body, in document order. */
export function readNoteReferences(doc: Node): ReadNoteReference[] {
	const references: ReadNoteReference[] = [];
	doc.descendants((node, pos) => {
		if (node.type.name !== 'noteReference') return true;
		references.push({
			pos,
			kind: node.attrs.kind === 'endnote' ? 'endnote' : 'footnote',
			id: String(node.attrs.id ?? ''),
			label: String(node.attrs.label ?? node.attrs.number),
		});
		return false;
	});
	return references;
}

export function findReadNote(model: DocumentModel, reference: ReadNoteReference): Note | undefined {
	const notes = reference.kind === 'endnote' ? model.endnotes : model.footnotes;
	return notes?.find((note) => note.id === reference.id);
}

export const noteTitle = (locale: EditorLocale, reference: ReadNoteReference): string =>
	translateTemplate(
		locale,
		reference.kind === 'endnote' ? 'readMode.endnote' : 'readMode.footnote',
		{
			number: reference.label,
		},
	);

const readNotesKey = new PluginKey<DecorationSet>('dve-read-notes');

/**
 * Read Mode shows a note in a callout when its reference is clicked or activated with Enter or
 * Space, as Word's Read Mode does; Escape or Back to text closes it and returns to the reference.
 */
export function mountReadNotes(
	frame: HTMLElement,
	locale: () => EditorLocale,
	model: () => DocumentModel,
) {
	let view: EditorView | undefined;
	let decoratedLocale: EditorLocale | undefined;
	let open: { reference: ReadNoteReference; callout: HTMLElement } | undefined;
	const plugin = new Plugin<DecorationSet>({
		key: readNotesKey,
		state: {
			init: (_, state) => decorate(state.doc),
			apply: (tr, set) =>
				tr.docChanged || tr.getMeta(readNotesKey) ? decorate(tr.doc) : set.map(tr.mapping, tr.doc),
		},
		props: { decorations: (state) => readNotesKey.getState(state) },
	});
	function decorate(doc: Node): DecorationSet {
		const current = locale();
		return DecorationSet.create(
			doc,
			readNoteReferences(doc)
				.filter((reference) => findReadNote(model(), reference))
				.map((reference) =>
					Decoration.node(reference.pos, reference.pos + 1, {
						tabindex: '0',
						role: 'button',
						'aria-haspopup': 'dialog',
						'aria-label': noteTitle(current, reference),
						'data-read-note': String(reference.pos),
					}),
				),
		);
	}
	const referenceAt = (target: EventTarget | null): ReadNoteReference | undefined => {
		const element =
			target instanceof Element ? target.closest<HTMLElement>('[data-read-note]') : null;
		if (!view || !element || !view.dom.contains(element)) return undefined;
		const pos = Number(element.dataset.readNote);
		return readNoteReferences(view.state.doc).find((reference) => reference.pos === pos);
	};
	const referenceElement = (reference: ReadNoteReference) => {
		const node = view?.nodeDOM(reference.pos);
		return node instanceof HTMLElement ? node : undefined;
	};

	function close(returnFocus: boolean) {
		if (!open) return;
		const { reference, callout } = open;
		open = undefined;
		callout.remove();
		if (returnFocus) referenceElement(reference)?.focus({ preventScroll: true });
	}
	function show(reference: ReadNoteReference) {
		const note = findReadNote(model(), reference);
		if (!note) return;
		close(false);
		const doc = frame.ownerDocument;
		const callout = doc.createElement('div');
		callout.className = 'dve-reader-note';
		callout.setAttribute('role', 'dialog');
		const title = doc.createElement('div');
		title.className = 'dve-reader-note-title';
		title.id = `dve-reader-note-${reference.kind}-${reference.id}`;
		title.textContent = noteTitle(locale(), reference);
		callout.setAttribute('aria-labelledby', title.id);
		const body = doc.createElement('div');
		body.className = 'dve-reader-note-body';
		body.append(renderStoryPreview(note.blocks, model()));
		const back = doc.createElement('button');
		back.type = 'button';
		back.className = 'dve-reader-note-back';
		back.textContent = translate(locale(), 'readMode.backToText');
		back.addEventListener('click', () => close(true));
		callout.append(title, body, back);
		callout.addEventListener('keydown', (event) => {
			if (event.key !== 'Escape') return;
			event.preventDefault();
			event.stopPropagation();
			close(true);
		});
		callout.addEventListener('focusout', (event) => {
			const next = event.relatedTarget;
			if (next instanceof globalThis.Node && !callout.contains(next)) close(false);
		});
		frame.append(callout);
		open = { reference, callout };
		place(callout, referenceElement(reference));
		back.focus({ preventScroll: true });
	}
	function place(callout: HTMLElement, anchor: HTMLElement | undefined) {
		const win = frame.ownerDocument.defaultView;
		const rect = anchor?.getBoundingClientRect();
		if (!rect || !win) return;
		const width = Math.min(360, win.innerWidth - 32);
		callout.style.width = `${width}px`;
		callout.style.left = `${Math.max(16, Math.min(rect.left - width / 2, win.innerWidth - width - 16))}px`;
		const below = rect.bottom + 8;
		const height = callout.offsetHeight;
		callout.style.top = `${below + height > win.innerHeight - 16 ? Math.max(16, rect.top - height - 8) : below}px`;
	}

	const click = (event: MouseEvent) => {
		const reference = referenceAt(event.target);
		if (reference) {
			event.preventDefault();
			show(reference);
		} else if (open && !open.callout.contains(event.target as globalThis.Node)) close(false);
	};
	frame.addEventListener('click', click);

	return {
		/** Closes an open callout and returns to its reference; true when one was open. */
		escape(): boolean {
			if (!open) return false;
			close(true);
			return true;
		},
		/** Handles Enter and Space on a focused reference before Read Mode pages with Space. */
		keydown(event: KeyboardEvent): boolean {
			if (event.key !== 'Enter' && event.key !== ' ') return false;
			const reference = referenceAt(event.target);
			if (!reference) return false;
			event.preventDefault();
			event.stopPropagation();
			show(reference);
			return true;
		},
		sync(next: EditorView | undefined) {
			if (next !== view) {
				uninstall();
				view = next && !next.isDestroyed ? next : undefined;
				if (view) {
					decoratedLocale = locale();
					view.updateState(view.state.reconfigure({ plugins: [plugin, ...view.state.plugins] }));
				}
			} else if (view && decoratedLocale !== locale()) {
				decoratedLocale = locale();
				view.dispatch(view.state.tr.setMeta(readNotesKey, true).setMeta('addToHistory', false));
			}
		},
		dispose() {
			close(false);
			frame.removeEventListener('click', click);
			uninstall();
			view = undefined;
		},
	};
	function uninstall() {
		if (view && !view.isDestroyed)
			view.updateState(
				view.state.reconfigure({
					plugins: view.state.plugins.filter((installed) => installed !== plugin),
				}),
			);
	}
}
