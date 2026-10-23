/**
 * viewer-keyboard.service.ts: Viewer-scoped editing keyboard-shortcut handler.
 *
 * Key-to-action resolution and the dispatch switch are the shared
 * `dispatchEditorKey`, the one keymap and handler interface every binding uses,
 * so Angular cannot drift from React / Vue / Vanilla / Svelte again. This
 * service only supplies the guard state and the `EditorKeyHandlers` that act on
 * {@link EditorStateService}.
 *
 * F5 / Shift+F5 (real PowerPoint's "start the show") are checked by the same
 * dispatcher ahead of the editor gates, because they must fire regardless of
 * `canEdit` or the event target: PowerPoint starts the show even from a
 * read-only deck and even with the caret in a text box.
 *
 * Extracted from {@link PowerPointViewerComponent}: the component keeps the thin
 * `@HostListener('document:keydown')` (a decorator can only live on the
 * component) and forwards the event to {@link handleKeyDown}; the host binds the
 * canEdit / presenting / active-slide-index / navigation accessors via
 * {@link bind}.
 *
 * Provide it once on the viewer component (`providers: [ViewerKeyboardService]`).
 */

import { inject, Injectable } from '@angular/core';
import {
	cycleSelectableElement,
	dispatchEditorKey,
	EMPTY_RESOLVED_CUSTOMIZATION,
	isFeatureEnabled,
} from 'ooxml-ui/pptx';
import type { EditorKeyHandlers } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';

import { EditorStateService } from './editor-state.service';
import { escapeSelectionTarget } from './group-drill-canvas';
import { PictureCropService } from './picture-crop.service';
import { ViewerCanvasEditingService } from './viewer-canvas-editing.service';
import { ViewerCustomizationService } from './viewer-customization.service';
import { ViewerDialogsService } from './viewer-dialogs.service';
import { ViewerDocumentPropertiesService } from './viewer-document-properties.service';
import { ViewerFindReplaceService } from './viewer-find-replace.service';
import { ViewerFormatPainterService } from './viewer-format-painter.service';
import { applyTextCommand, stepSelectionFontSize } from './viewer-keyboard-text-commands';
import { ViewerPresentationModeService } from './viewer-presentation-mode.service';

/** Live host accessors the shortcut handler consults. */
interface KeyboardHost {
	readonly canEdit: () => boolean;
	readonly presenting: () => boolean;
	readonly activeSlideIndex: () => number;
	/**
	 * A Draw-tab tool other than the selection arrow is armed. The shared keymap
	 * stands its whole editing set down while one is, so a pen stroke is not also
	 * a Delete: every other binding passes this flag and Angular did not, which
	 * left Delete, Ctrl+D and the arrow nudges live over an armed pen.
	 */
	readonly isDrawing?: () => boolean;
	/** Go back one slide (the arrows page the deck when nothing is selected). */
	readonly goPrev?: () => void;
	/** Go forward one slide. */
	readonly goNext?: () => void;
	/**
	 * An inline text or table-cell editor is open. Needed so the text-command
	 * chords (alignment, font-size ladder, format painter, hyperlink, clear
	 * formatting) can survive `mapEditorKey`'s typing gate exactly as they do in
	 * the other four bindings, instead of only ever firing off a selection.
	 */
	readonly isEditingText?: () => boolean;
	/** The single selected element, the target of every text/format command. */
	readonly selectedElement?: () => PptxElement | null;
}

@Injectable()
export class ViewerKeyboardService {
	private readonly editor = inject(EditorStateService);
	private readonly dialogs = inject(ViewerDialogsService);
	private readonly formatPainter = inject(ViewerFormatPainterService);
	private readonly findReplace = inject(ViewerFindReplaceService);
	private readonly presentationMode = inject(ViewerPresentationModeService);
	private readonly docProperties = inject(ViewerDocumentPropertiesService);
	private readonly canvasEditing = inject(ViewerCanvasEditingService, { optional: true });
	private readonly customization = inject(ViewerCustomizationService, { optional: true });
	private readonly crop = inject(PictureCropService, { optional: true });

	private host: KeyboardHost | null = null;

	/** Wire the host accessors (called once from the component constructor). */
	bind(host: KeyboardHost): void {
		this.host = host;
	}

	handleKeyDown(event: KeyboardEvent): void {
		const host = this.host;
		if (!host) {
			return;
		}
		if (this.crop?.handleKeyDown(event)) {
			return;
		}
		const resolved = this.customization?.resolved() ?? EMPTY_RESOLVED_CUSTOMIZATION;
		dispatchEditorKey(
			event,
			{
				canEdit: host.canEdit(),
				canPaste: this.editor.hasClipboard(),
				isPresenting: host.presenting(),
				hasSelection: this.editor.hasSelection(),
				isDrawing: host.isDrawing?.() ?? false,
				isEditingText: host.isEditingText?.() ?? false,
				keyboard: resolved.keyboard,
				canStartShow: isFeatureEnabled(resolved, 'presentMode'),
			},
			this.keyHandlers(host),
		);
	}

	/** The shared editor key handlers, bound to the active slide of `host`. */
	private keyHandlers(host: KeyboardHost): EditorKeyHandlers {
		const idx = host.activeSlideIndex();
		const fontStep = (direction: 'increase' | 'decrease'): void =>
			stepSelectionFontSize(
				this.editor,
				idx,
				host.selectedElement?.() ?? null,
				direction,
				this.canvasEditing,
			);
		return {
			startShowFromBeginning: () => this.presentationMode.presentFromBeginning(),
			startShowFromCurrent: () => this.presentationMode.present(),
			escape: () => this.handleEscape(),
			toggleShortcuts: () => this.dialogs.showShortcuts.set(!this.dialogs.showShortcuts()),
			find: () => this.toggleFind(),
			undo: () => this.editor.undo(),
			redo: () => this.editor.redo(),
			duplicate: () => this.editor.duplicateSelected(idx),
			copy: () => this.editor.copySelected(idx),
			cut: () => this.editor.cutSelected(idx),
			paste: () => this.editor.paste(idx),
			selectAll: () => this.editor.selectAll(idx),
			group: () => this.editor.groupSelected(idx),
			ungroup: () => this.editor.ungroupSelected(idx),
			delete: () => this.editor.deleteSelected(idx),
			nudge: (dx, dy) => this.editor.moveSelectedBy(idx, dx, dy),
			prevSlide: () => host.goPrev?.(),
			nextSlide: () => host.goNext?.(),
			alignLeft: () => this.applyTextPatch(idx, { align: 'left' }),
			alignCenter: () => this.applyTextPatch(idx, { align: 'center' }),
			alignRight: () => this.applyTextPatch(idx, { align: 'right' }),
			alignJustify: () => this.applyTextPatch(idx, { align: 'justify' }),
			increaseFontSize: () => fontStep('increase'),
			decreaseFontSize: () => fontStep('decrease'),
			copyFormat: () => this.formatPainter.toggle(),
			pasteFormat: () => this.pasteFormat(),
			newSlide: () => this.editor.addSlide(idx),
			hyperlink: () => this.openHyperlink(),
			findReplace: () => this.findReplace.openFindReplace(),
			clearFormatting: () =>
				this.applyTextPatch(idx, {
					bold: false,
					italic: false,
					underline: false,
					strikethrough: false,
				}),
			cycleSelectionNext: () => this.cycleSelection(idx, 'next'),
			cycleSelectionPrev: () => this.cycleSelection(idx, 'prev'),
			pasteSpecial: () => this.editor.isPasteSpecialDialogOpen.set(true),
		};
	}

	private applyTextPatch(slideIndex: number, patch: Parameters<typeof applyTextCommand>[3]): void {
		applyTextCommand(
			this.editor,
			slideIndex,
			this.host?.selectedElement?.() ?? null,
			patch,
			this.canvasEditing,
		);
	}

	/** Ctrl+Shift+V: apply the copied format to the current selection (no click needed). */
	private pasteFormat(): void {
		const id = this.editor.selectedIds()[0];
		if (id) {
			this.formatPainter.applyToTarget(id);
		}
	}

	/** Ctrl+K, gated on a selection exactly like the ribbon's hyperlink button. */
	private openHyperlink(): void {
		if (this.editor.hasSelection()) {
			this.docProperties.showHyperlink.set(true);
		}
	}

	/** Tab / Shift+Tab: move the selection to the next/previous element in z-order. */
	private cycleSelection(slideIndex: number, direction: 'next' | 'prev'): void {
		const ids = this.editor.slides()[slideIndex]?.elements.map((el) => el.id) ?? [];
		const nextId = cycleSelectableElement(ids, this.editor.selectedIds()[0] ?? null, direction);
		if (nextId) {
			this.editor.select([nextId]);
		}
	}

	/**
	 * Ctrl/Cmd+F toggles the find bar. Angular has shipped the bar since the
	 * find-replace port but had no shortcut for it at all, because the chord was
	 * hand-wired in React and Vue instead of living in the shared keymap.
	 *
	 * The full find-and-replace bar counts as "open" for the toggle: pressing the
	 * chord while it is up closes it rather than swapping it for the smaller find
	 * bar, which is what every other binding's single-panel toggle does.
	 */
	private toggleFind(): void {
		if (this.findReplace.showFind() || this.findReplace.showFindReplace()) {
			this.findReplace.showFind.set(false);
			this.findReplace.closeFindReplace();
			return;
		}
		this.findReplace.showFind.set(true);
	}

	/**
	 * Escape unwinds the transient chrome one layer at a time: an armed format
	 * painter first (it is modal over the pointer), then the shortcut cheat
	 * sheet, then a selected group member steps up to its group, then the
	 * selection itself (as in React, Svelte and Vanilla, so the selection's
	 * contextual ribbon tabs go away too). Before this the branch
	 * stopped after the painter, so the panel "?" had just opened stayed on
	 * screen with no key that could dismiss it.
	 */
	private handleEscape(): void {
		if (this.formatPainter.active()) {
			this.formatPainter.cancel();
			return;
		}
		if (this.dialogs.showShortcuts()) {
			this.dialogs.showShortcuts.set(false);
			return;
		}
		// A member selected inside a group steps back out to its group first
		// (shared `parentSelection`); a top-level selection clears.
		const slide = this.editor.slides()[this.host?.activeSlideIndex() ?? -1];
		const parent = escapeSelectionTarget(slide?.elements, this.editor.selectedIds());
		if (parent) {
			this.editor.select([parent]);
			return;
		}
		this.editor.clearSelection();
	}
}
