import type { Block, DocumentModel, Note, NoteSelection, ReviewDisplayMode } from 'ooxml-core/docx';
import type { FindStory } from 'ooxml-core/docx/ui';
import { TextSelection, type Plugin } from 'prosemirror-state';
import { closeHistory, undo, redo } from 'prosemirror-history';
import { runStylesPlugin } from './run-styles';
import { paragraphStylesPlugin } from './paragraph-styles';
import { reviewDisplayPlugin } from './review-display';
import { reviewRunMarksPlugin } from './review-run-marks';
import type { EditorView } from 'prosemirror-view';
import type { EditorHost } from './editor-host';
import { buildFooterElement, buildHeaderElement } from './header-footer-view';
import { refreshStoryPreviews } from './review-story-previews';
import {
	attachHeaderFooterEditing,
	openBlocksEditor,
	type HeaderFooterSlotName,
	type InlineEditorOptions,
} from './header-footer-editor';
import { imageNodeView, type ImageMediaCache } from './image-media';
import { buildNotesElement } from './notes-view';
import { attachNoteEditing, withMark, withoutMark } from './note-editor';
import { convertNotesInView, insertNote, type NoteKind } from './note-commands';
import { sectionPartsJson, HEADER_FOOTER_INPUT } from 'ooxml-core/docx/ui';
import { effectiveHeaderFooter, withHeaderFooterLink } from './header-footer-link';
import type { HeaderFooterContext } from './header-footer-ribbon';
import { storyPreview, selectSectionStart } from './header-footer-navigation';
import {
	headerFooterStoryKey,
	headerFooterTransaction,
	notesTransaction,
} from './parts-story-writes';
import type { HeaderFooterStory } from './search-story-replace';

export interface PartsControllerHost extends EditorHost {
	images(): ImageMediaCache;
	/** Plugins shared with in-place editors (e.g. Ctrl+K). */
	plugins(): Plugin[];
	/** Where focus may go without closing an in-place editor: the ribbon and dialogs. */
	keepOpenWithin(): (Element | undefined)[];
	sectionIndex?(): number;
	refreshControls?(): void;
	reviewDisplayMode?(): ReviewDisplayMode;
	/** The model as last loaded or saved, whose note ids a conversion must not reuse. */
	savedModel?(): DocumentModel | undefined;
}

/** Headers, footers, footnotes and endnotes around the continuous editing surface. */
export class PartsController {
	private headerEl?: HTMLElement | undefined;
	private footerEl?: HTMLElement | undefined;
	private notesEl?: HTMLElement | undefined;
	private active?: EditorView | undefined;
	private activeNote?: { kind: NoteKind; id: string } | undefined;
	private closeActive?: (() => void) | undefined;
	private headerContext?: { kind: 'headers' | 'footers'; slot: HeaderFooterSlotName } | undefined;
	private shownSection = 0;
	private renderingContext?: HeaderFooterContext | undefined;

	constructor(private readonly host: PartsControllerHost) {}
	private readonly reviewMode = (): ReviewDisplayMode => this.host.reviewDisplayMode?.() ?? 'all';
	/** Story previews show pictures from the body's media cache, so they share its URLs. */
	private readonly pictureUrl = (partName: string, contentType: string) =>
		this.host.images().urlFor(partName, contentType);

	/** Refresh appearance in place, retaining the active story's selection and history. */
	refreshReviewDisplay(): void {
		this.active?.dispatch(this.active.state.tr.setMeta('addToHistory', false));
		refreshStoryPreviews(
			this.host.model(),
			this.shownSection,
			{ headers: this.headerEl, footers: this.footerEl, notes: this.notesEl },
			this.reviewMode(),
			this.pictureUrl,
		);
	}

	/** The in-place header/footer/note editor currently open, which the ribbon then targets. */
	activeView(): EditorView | undefined {
		return this.active;
	}
	/** Page-number insertion follows the first/even/default story currently being edited. */
	activeHeaderFooterSlot(): HeaderFooterSlotName | undefined {
		return this.headerContext?.slot;
	}
	headerFooterContext(): HeaderFooterContext | undefined {
		return this.headerContext
			? { ...this.headerContext, index: this.shownSection }
			: this.renderingContext;
	}
	/**
	 * Edits a header or footer in place on a Print Layout page: `container` is that page's element.
	 * The ribbon then targets it (Header & Footer tab) until it closes, which calls `closed`.
	 */
	editOnPage(story: {
		container: HTMLElement;
		kind: 'headers' | 'footers';
		slot: HeaderFooterSlotName;
		sectionIndex: number;
		blocks: Block[];
		closed(): void;
	}): boolean {
		if (!this.host.view()?.editable || !this.host.canEditOutsideBody()) return false;
		this.closeActive?.();
		this.shownSection = story.sectionIndex;
		const base = this.editorOptions(true);
		openBlocksEditor(
			story.container,
			story.container,
			story.blocks,
			(blocks) => this.updateHeaderFooter(story.kind, story.slot, blocks),
			{
				...base,
				activate: (view) => {
					base.activate?.(view);
					this.headerContext = { kind: story.kind, slot: story.slot };
					this.host.refreshControls?.();
				},
				deactivate: (view) => {
					base.deactivate?.(view);
					story.closed();
				},
			},
		);
		return true;
	}
	/**
	 * Edits a footnote or endnote in place on a Print Layout page: `container` is the element the
	 * note is drawn in. The ribbon targets the note until it closes (Escape, or focus leaving), which
	 * calls `closed`. `opened` receives the editor so the caller can place the caret.
	 */
	editNoteOnPage(story: {
		container: HTMLElement;
		kind: NoteKind;
		id: string;
		opened?(view: EditorView): void;
		closed(): void;
	}): boolean {
		if (!this.host.view()?.editable || !this.host.canEditOutsideBody()) return false;
		const model = this.host.model();
		const notes = story.kind === 'footnote' ? model.footnotes : model.endnotes;
		const note = notes?.find((item) => item.id === story.id);
		if (!note) return false;
		this.closeActive?.();
		const { editable, mark } = withoutMark(note.blocks);
		const base = this.editorOptions();
		openBlocksEditor(
			story.container,
			story.container,
			editable,
			(blocks) => this.updateNote(story.kind, story.id, withMark(blocks, mark)),
			{
				...base,
				activate: (view) => {
					base.activate?.(view);
					this.activeNote = { kind: story.kind, id: story.id };
					story.opened?.(view);
				},
				deactivate: (view) => {
					base.deactivate?.(view);
					story.closed();
				},
			},
		);
		return true;
	}
	/** Closes the open header, footer or note editor, leaving the body focused unless told not to. */
	closeOpenStory(focusBody = true): void {
		this.closeActive?.();
		if (focusBody) this.host.view()?.focus();
	}
	closeHeaderFooter(canvas?: HTMLElement, paper?: HTMLElement): void {
		this.closeActive?.();
		this.render(canvas, paper);
		this.host.view()?.focus();
	}
	navigate(
		target: 'header' | 'footer' | 'previous' | 'next',
		canvas?: HTMLElement,
		paper?: HTMLElement,
	): void {
		const context = this.headerContext;
		const view = this.host.view();
		if (!context || !view) return;
		const index = this.shownSection + (target === 'previous' ? -1 : target === 'next' ? 1 : 0);
		if (index < 0 || index >= (this.host.model().sections?.length ?? 0)) return;
		this.shownSection = index;
		this.headerContext = {
			...context,
			kind: target === 'header' ? 'headers' : target === 'footer' ? 'footers' : context.kind,
		};
		selectSectionStart(view, this.host.model(), index);
		this.render(canvas, paper, true);
	}
	toggleLink(): void {
		const context = this.headerContext;
		const view = this.host.view();
		if (!context || !view?.editable || !this.host.canEditOutsideBody()) return;
		const model = this.host.model();
		const linked = !model.sections?.[this.shownSection]?.[context.kind]?.[context.slot];
		const next = withHeaderFooterLink(
			model,
			this.shownSection,
			context.kind,
			context.slot,
			!linked,
		);
		if (next !== model)
			view.dispatch(
				closeHistory(view.state.tr).setDocAttribute(
					'sectionParts',
					sectionPartsJson(next.sections),
				),
			);
	}

	private editorOptions(header = false): InlineEditorOptions {
		return {
			contextModel: () => this.host.model(),
			reviewDisplayMode: this.reviewMode,
			pictureUrl: this.pictureUrl,
			nodeViews: {
				image: imageNodeView(this.host.images(), {
					theme: () => this.host.model().theme,
					model: () => this.host.model(),
				}),
			},
			plugins: [
				...this.host.plugins(),
				reviewDisplayPlugin(this.reviewMode),
				reviewRunMarksPlugin(this.reviewMode),
				runStylesPlugin(() => this.host.model(), this.reviewMode),
				paragraphStylesPlugin(
					() => this.host.model(),
					this.reviewMode,
					(partName, contentType) => this.host.images().urlFor(partName, contentType),
				),
			],
			keepOpenWithin: () => this.host.keepOpenWithin(),
			activate: (view) => {
				this.active = view;
				if (header) this.host.view()?.dispatch(closeHistory(this.host.view()!.state.tr));
			},
			deactivate: (view) => {
				if (this.active === view) {
					this.active = undefined;
					this.activeNote = undefined;
					this.closeActive = undefined;
				}
				if (header) {
					this.headerContext = undefined;
					this.host.view()?.dispatch(closeHistory(this.host.view()!.state.tr));
					this.host.refreshControls?.();
				}
			},
			registerClose: (close) => {
				this.closeActive = close;
			},
			...(header ? { history: (key: 'undo' | 'redo') => this.runHistory(key) } : {}),
		};
	}

	/** Header/footer typing and insertion use the body history; notes retain their local stack. */
	runHistory(key: 'undo' | 'redo'): boolean {
		const view = this.host.view();
		return !!view?.editable && (key === 'undo' ? undo : redo)(view.state, view.dispatch);
	}
	usesBodyHistory(): boolean {
		return !!this.headerContext;
	}

	syncSection(canvas?: HTMLElement, paper?: HTMLElement): void {
		if (!this.active && this.shownSection !== (this.host.sectionIndex?.() ?? 0))
			this.render(canvas, paper);
	}

	/** Rebuilds the header/footer/note previews and wires in-place editing. */
	render(
		canvas: HTMLElement | undefined,
		paper: HTMLElement | undefined,
		keepEditor = false,
	): void {
		if (!canvas || !paper) return;
		const context = keepEditor ? this.headerContext : undefined;
		const selection = context ? this.active?.state.selection : undefined;
		if (!context) this.shownSection = this.host.sectionIndex?.() ?? 0;
		this.renderingContext = context ? { ...context, index: this.shownSection } : undefined;
		this.closeActive?.();
		this.headerEl?.remove();
		this.footerEl?.remove();
		this.notesEl?.remove();
		const model = storyPreview(this.host.model(), this.renderingContext);
		const locale = this.host.locale();
		this.headerEl =
			buildHeaderElement(model, locale, this.shownSection, this.pictureUrl) ?? undefined;
		this.footerEl =
			buildFooterElement(model, locale, this.shownSection, this.pictureUrl) ?? undefined;
		this.notesEl = buildNotesElement(model, locale, this.pictureUrl) ?? undefined;
		const editable = () => this.host.canEditOutsideBody();
		for (const [element, kind] of [
			[this.headerEl, 'headers'],
			[this.footerEl, 'footers'],
		] as const)
			if (element)
				attachHeaderFooterEditing(element, {
					content: (slot) =>
						effectiveHeaderFooter(this.host.model(), this.shownSection, kind, slot) ??
						effectiveHeaderFooter(model, this.shownSection, kind, slot),
					change: (slot, blocks) => this.updateHeaderFooter(kind, slot, blocks),
					editable,
					editor: {
						...this.editorOptions(true),
						activate: (view) => {
							this.editorOptions(true).activate?.(view);
							const slot = view.dom.closest<HTMLElement>('[data-slot]')?.dataset
								.slot as HeaderFooterSlotName;
							this.headerContext = { kind, slot };
							this.host.refreshControls?.();
						},
					},
				});
		if (this.notesEl)
			attachNoteEditing(this.notesEl, {
				note: (kind, id) => {
					const current = this.host.model();
					return current[kind === 'footnote' ? 'footnotes' : 'endnotes']?.find(
						(note) => note.id === id,
					);
				},
				change: (kind, id, blocks) => this.updateNote(kind, id, blocks),
				editable,
				editor: {
					...this.editorOptions(),
					activate: (view) => {
						this.editorOptions().activate?.(view);
						const item = view.dom.closest<HTMLElement>('li[data-docx-note-id]');
						const kind = item?.dataset.docxNoteKind;
						if (item && (kind === 'footnote' || kind === 'endnote'))
							this.activeNote = { kind, id: item.dataset.docxNoteId ?? '' };
					},
				},
			});
		if (this.headerEl) canvas.insertBefore(this.headerEl, paper);
		if (this.footerEl) canvas.insertBefore(this.footerEl, paper.nextSibling);
		if (this.notesEl) canvas.insertBefore(this.notesEl, (this.footerEl ?? paper).nextSibling);
		if (context) {
			const root = context.kind === 'headers' ? this.headerEl : this.footerEl;
			root
				?.querySelector<HTMLElement>(`[data-slot="${context.slot}"]`)
				?.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
			if (this.active && selection) {
				const size = this.active.state.doc.content.size;
				this.active.dispatch(
					this.active.state.tr.setSelection(
						TextSelection.near(this.active.state.doc.resolve(Math.min(selection.from, size))),
					),
				);
			} else this.host.view()?.focus();
		}
		this.refreshReviewDisplay();
		this.renderingContext = undefined;
		this.host.refreshControls?.();
	}

	/** Inserts a footnote or endnote reference at the selection and opens the new note for typing. */
	insertNote(kind: NoteKind, canvas?: HTMLElement, paper?: HTMLElement): void {
		const view = this.host.view();
		if (!view?.editable || !this.host.canEditOutsideBody()) return;
		const { id } = insertNote(view, this.host.model(), kind);
		this.render(canvas, paper);
		const item = this.notesEl?.querySelector<HTMLElement>(
			`.dve-notes-${kind} li[data-docx-note-id="${id}"]`,
		);
		item?.scrollIntoView?.({ block: 'center' });
		item?.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
		this.host.edited();
	}

	/** References > Convert Notes, or Convert to Endnote/Footnote on one note's reference. */
	convertNotes(selection: NoteSelection): void {
		const view = this.host.view();
		if (!view?.editable || !this.host.canEditOutsideBody()) return;
		this.closeActive?.();
		const saved = this.host.savedModel?.();
		const ids = (notes: Note[] | undefined) => (notes ?? []).map((note) => note.id);
		const plan = convertNotesInView(view, this.host.model(), selection, {
			footnote: ids(saved?.footnotes),
			endnote: ids(saved?.endnotes),
		});
		if (plan.moved.size) this.host.edited();
		if (plan.skipped.length)
			this.host.reportError(
				new Error(
					`${plan.skipped.length} note(s) hold charts, shapes or floating pictures and were not converted.`,
				),
			);
	}

	/** Replaces one footnote's or endnote's blocks. */
	private updateNote(kind: NoteKind, id: string, blocks: Block[]): void {
		const view = this.host.view();
		if (!view?.editable || !this.host.canEditOutsideBody()) return;
		const tr = notesTransaction(view.state, this.host.model(), kind, new Map([[id, blocks]]));
		view.dispatch(tr.setMeta(HEADER_FOOTER_INPUT, true));
	}

	/** Applies header/footer edits to every section slot that shares the edited part. */
	private updateHeaderFooter(
		kind: 'headers' | 'footers',
		slot: HeaderFooterSlotName,
		blocks: Block[],
	): void {
		const view = this.host.view();
		if (!view?.editable || !this.host.canEditOutsideBody()) return;
		const model = this.host.model();
		const tr = headerFooterTransaction(view.state, model, kind, slot, this.shownSection, blocks);
		view.dispatch(tr.setMeta(HEADER_FOOTER_INPUT, true));
	}

	/** The story key (`findStoryKey`) of the open header, footer or note editor. */
	activeStoryKey(): string | undefined {
		if (!this.active) return undefined;
		const context = this.headerContext;
		if (context)
			return headerFooterStoryKey(this.host.model(), context.kind, context.slot, this.shownSection);
		return this.activeNote && `${this.activeNote.kind}:${this.activeNote.id}`;
	}

	/**
	 * Find's way into a story on the continuous surface: shows the header or footer's section (or
	 * the note), opens its editor when editing is allowed and selects `from`..`to` in it.
	 */
	showStory(
		story: FindStory,
		from: number,
		to: number,
		canvas?: HTMLElement,
		paper?: HTMLElement,
	): void {
		const view = this.host.view();
		if (!view) return;
		this.closeActive?.();
		let target: HTMLElement | null | undefined;
		if (story.kind === 'header' || story.kind === 'footer') {
			if (this.shownSection !== story.sectionIndex || !this.headerEl)
				selectSectionStart(view, this.host.model(), story.sectionIndex);
			this.render(canvas, paper);
			const root = story.kind === 'header' ? this.headerEl : this.footerEl;
			target = root?.querySelector<HTMLElement>(`[data-slot="${story.slot}"]`);
		} else if (story.kind === 'footnote' || story.kind === 'endnote')
			target = [
				...(this.notesEl?.querySelectorAll<HTMLElement>(
					`li[data-docx-note-kind="${story.kind}"]`,
				) ?? []),
			].find((item) => item.dataset.docxNoteId === story.id);
		target?.scrollIntoView?.({ block: 'center' });
		target?.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
		const active = this.active;
		if (!active) return;
		const size = active.state.doc.content.size;
		const selection = TextSelection.create(
			active.state.doc,
			Math.min(from, size),
			Math.min(to, size),
		);
		active.dispatch(active.state.tr.setSelection(selection).scrollIntoView());
	}

	/** Replace All's write of a header or footer part, as its own undo step. */
	writeHeaderFooter(story: HeaderFooterStory, blocks: Block[]): void {
		const view = this.host.view();
		if (!view?.editable || !this.host.canEditOutsideBody()) return;
		const kind = story.kind === 'header' ? 'headers' : 'footers';
		const model = this.host.model();
		view.dispatch(
			closeHistory(
				headerFooterTransaction(view.state, model, kind, story.slot, story.sectionIndex, blocks),
			),
		);
	}

	/** Replace All's write of the footnotes or endnotes story, as one undo step. */
	writeNotes(kind: NoteKind, blocks: ReadonlyMap<string, Block[]>): void {
		const view = this.host.view();
		if (!view?.editable || !this.host.canEditOutsideBody() || !blocks.size) return;
		view.dispatch(closeHistory(notesTransaction(view.state, this.host.model(), kind, blocks)));
	}
}
