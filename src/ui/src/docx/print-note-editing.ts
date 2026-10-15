import type { Block, DocumentModel, Note } from 'ooxml-core/docx';
import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import type { NoteKind } from './note-commands';
import type { PartsController } from './parts-controller';

export interface PrintNoteHost {
	/** The `.dve-print-pages` element. */
	pages: HTMLElement;
	parts: Pick<PartsController, 'editNoteOnPage' | 'closeOpenStory'>;
	model(): DocumentModel;
	/** The body editor, which keeps the caret outside the note. */
	view(): EditorView | undefined;
	printActive(): boolean;
	holdLayout(held: boolean): void;
	/** Mounts the page that holds a paragraph (virtualised documents). */
	showBlock(id: string, offset: number): void;
}

interface NoteRef {
	kind: NoteKind;
	id: string;
	/** Id of the note's first paragraph (where the automatic number mark sits). */
	firstId: string;
}

const EDITING = 'dve-print-note-editing';

function paragraphIds(blocks: Block[]): string[] {
	return blocks.flatMap((block) =>
		block.type === 'paragraph'
			? [block.id]
			: block.rows.flatMap((row) => row.flatMap((cell) => paragraphIds(cell.paragraphs))),
	);
}

/** Paragraph id to owning note, for every footnote and endnote paragraph in the model. */
export function noteParagraphs(model: DocumentModel): Map<string, NoteRef> {
	const map = new Map<string, NoteRef>();
	const add = (kind: NoteKind, notes: Note[] | undefined) => {
		for (const note of notes ?? []) {
			const ids = paragraphIds(note.blocks);
			for (const id of ids) map.set(id, { kind, id: note.id, firstId: ids[0] ?? '' });
		}
	};
	add('footnote', model.footnotes);
	add('endnote', model.endnotes);
	return map;
}

/** The note a body reference mark under `target` points at, when it is one. */
export function referenceUnder(
	view: EditorView,
	target: Element,
): { kind: NoteKind; id: string } | undefined {
	const fragment = target.closest<HTMLElement>('span[data-source-start]');
	const line = fragment?.closest<HTMLElement>('.dve-print-line[data-block-id]');
	if (!fragment || !line || line.closest('.dve-print-footnotes')) return undefined;
	const start = Number(fragment.dataset.sourceStart);
	let found: { kind: NoteKind; id: string } | undefined;
	view.state.doc.descendants((node) => {
		if (found) return false;
		if (node.type.name !== 'paragraph') return true;
		if (String(node.attrs.id) === line.dataset.blockId)
			node.forEach((child, offset) => {
				if (offset === start && child.type.name === 'noteReference')
					found = {
						kind: child.attrs.kind === 'endnote' ? 'endnote' : 'footnote',
						id: child.attrs.id,
					};
			});
		return false;
	});
	return found;
}

/** Position just after the body reference to a note, with the id of its paragraph. */
export function referencePosition(
	view: EditorView,
	ref: { kind: NoteKind; id: string },
): { pos: number; blockId: string; offset: number } | undefined {
	let found: { pos: number; blockId: string; offset: number } | undefined;
	view.state.doc.descendants((node, pos) => {
		if (found) return false;
		if (node.type.name !== 'paragraph') return true;
		node.forEach((child, offset) => {
			if (
				!found &&
				child.type.name === 'noteReference' &&
				child.attrs.id === ref.id &&
				(child.attrs.kind === 'endnote' ? 'endnote' : 'footnote') === ref.kind
			)
				found = {
					pos: pos + 1 + offset + child.nodeSize,
					blockId: String(node.attrs.id),
					offset: offset + child.nodeSize,
				};
		});
		return false;
	});
	return found;
}

/**
 * Notes on Print Layout pages. A press on a note's text opens the note's own editor over it with
 * the caret there; Escape, or a press anywhere else, returns to the body. Double-clicking a
 * reference mark in the body jumps to its note, and double-clicking the number in the note jumps
 * back to the reference, as in Word.
 */
export function attachPrintNoteEditing(host: PrintNoteHost): () => void {
	const controller = new AbortController();
	let editing: { container: HTMLElement; hidden: HTMLElement[] } | undefined;

	const noteElements = (
		ref: NoteRef,
	): { container: HTMLElement; rest: HTMLElement[] } | undefined => {
		if (ref.kind === 'footnote') {
			const all = [
				...host.pages.querySelectorAll<HTMLElement>('.dve-print-footnote[data-note-id]'),
			].filter((element) => element.dataset.noteId === ref.id);
			const container = all.find((element) => !element.dataset.continued) ?? all[0];
			return container
				? { container, rest: all.filter((element) => element !== container) }
				: undefined;
		}
		const ids = new Set(
			[...noteParagraphs(host.model())].filter(([, note]) => note.id === ref.id).map(([id]) => id),
		);
		const blocks = [...host.pages.querySelectorAll<HTMLElement>('.dve-print-line[data-block-id]')]
			.filter((line) => ids.has(line.dataset.blockId!))
			.map((line) => line.closest<HTMLElement>('.dve-print-block'))
			.filter(
				(block, index, list): block is HTMLElement => !!block && list.indexOf(block) === index,
			);
		const [container, ...rest] = blocks;
		return container ? { container, rest } : undefined;
	};

	const finish = () => {
		if (!editing) return;
		for (const element of editing.hidden) element.style.visibility = '';
		editing = undefined;
		host.pages.classList.remove(EDITING);
		host.holdLayout(false);
		host.view()?.focus();
	};

	const open = (ref: NoteRef, at?: { x: number; y: number }): boolean => {
		host.showBlock(ref.firstId, 0);
		const elements = noteElements(ref);
		if (!elements) return false;
		const { container, rest } = elements;
		const mark = container.querySelector<HTMLElement>('.dve-print-line span');
		const label = mark?.textContent ?? '';
		host.holdLayout(true);
		for (const element of rest) element.style.visibility = 'hidden';
		const opened = host.parts.editNoteOnPage({
			container,
			kind: ref.kind,
			id: ref.id,
			opened(view) {
				if (!at) {
					view.dispatch(view.state.tr.setSelection(TextSelection.atEnd(view.state.doc)));
					return;
				}
				const hit = view.posAtCoords({ left: at.x, top: at.y });
				view.dispatch(
					view.state.tr.setSelection(
						hit
							? TextSelection.near(view.state.doc.resolve(hit.pos))
							: TextSelection.atEnd(view.state.doc),
					),
				);
			},
			closed: finish,
		});
		if (!opened) {
			for (const element of rest) element.style.visibility = '';
			host.holdLayout(false);
			return false;
		}
		if (label && !container.querySelector('.dve-print-note-label')) {
			const sup = document.createElement('sup');
			sup.className = 'dve-print-note-label';
			sup.textContent = label;
			container.prepend(sup);
		}
		editing = { container, hidden: rest };
		container.classList.add('dve-print-note-container');
		host.pages.classList.add(EDITING);
		return true;
	};

	const noteAt = (target: Element): NoteRef | undefined => {
		const line = target.closest<HTMLElement>('.dve-print-line[data-block-id]');
		return line ? noteParagraphs(host.model()).get(line.dataset.blockId!) : undefined;
	};

	/** The automatic number at the start of a note (first fragment of its first paragraph). */
	const isMark = (target: Element, ref: NoteRef): boolean => {
		const fragment = target.closest<HTMLElement>('span[data-source-start]');
		const line = fragment?.closest<HTMLElement>('.dve-print-line[data-block-id]');
		if (!fragment || !line || line.dataset.blockId !== ref.firstId) return false;
		const note = (ref.kind === 'footnote' ? host.model().footnotes : host.model().endnotes)?.find(
			(item) => item.id === ref.id,
		);
		const first = note?.blocks[0];
		return (
			first?.type === 'paragraph' &&
			!!first.runs[0]?.noteMark &&
			line.dataset.sourceStart === '0' &&
			fragment.dataset.sourceStart === '0' &&
			fragment === line.querySelector('span[data-source-start]')
		);
	};

	host.pages.addEventListener(
		'pointerdown',
		(event) => {
			if (!host.printActive() || event.button !== 0 || !(event.target instanceof Element)) return;
			if (editing?.container.contains(event.target)) return;
			const ref = noteAt(event.target);
			const press = { x: event.clientX, y: event.clientY };
			const onMark = ref ? isMark(event.target, ref) : false;
			if (editing) host.parts.closeOpenStory();
			if (!ref || onMark) return;
			if (open(ref, press)) {
				event.preventDefault();
				event.stopImmediatePropagation();
			}
		},
		{ capture: true, signal: controller.signal },
	);

	host.pages.addEventListener(
		'dblclick',
		(event) => {
			if (!host.printActive() || !(event.target instanceof Element)) return;
			const view = host.view();
			if (!view) return;
			const ref = noteAt(event.target);
			if (ref) {
				if (!isMark(event.target, ref)) return;
				const reference = referencePosition(view, ref);
				if (!reference) return;
				if (editing) host.parts.closeOpenStory();
				host.showBlock(reference.blockId, reference.offset);
				view.dispatch(
					view.state.tr.setSelection(TextSelection.create(view.state.doc, reference.pos)),
				);
				view.focus();
				event.preventDefault();
				event.stopPropagation();
				return;
			}
			const target = referenceUnder(view, event.target);
			if (!target) return;
			const note = noteParagraphs(host.model());
			const first = [...note].find(
				([, item]) => item.kind === target.kind && item.id === target.id,
			)?.[1];
			if (!first) return;
			if (editing) host.parts.closeOpenStory();
			if (open(first)) {
				event.preventDefault();
				event.stopPropagation();
			}
		},
		{ capture: true, signal: controller.signal },
	);

	return () => {
		controller.abort();
		if (editing) host.parts.closeOpenStory();
	};
}
