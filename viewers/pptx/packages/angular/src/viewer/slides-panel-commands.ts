/**
 * slides-panel-commands.ts: the operations behind a host-driven slide
 * thumbnail rail (`SlidesPanelComponent` with `[slides]` set), applied to the
 * viewer's `EditorStateService`.
 *
 * A custom shell binds the rail's outputs to these, so a command from the
 * rail goes through the same editor operations, undo history and
 * collaboration broadcast as the rest of the shell. The rail's multi-selection
 * commands run through the shared `createSlidesPaneBulkSlideOps`, since the
 * editor only has single-slide operations: duplicates land right after their
 * own source and deletions never shift an index still to delete.
 *
 * @module angular-viewer/slides-panel-commands
 */
import { createSlidesPaneBulkSlideOps } from 'ooxml-ui/pptx';

import type { EditorStateService } from './editor-state.service';
import type { SlidesPanelMoveSlide, SlidesPanelRenameSection } from './slides-panel.component';

/** Handlers for the outputs of a host-driven `SlidesPanelComponent`, by output name. */
export interface SlidesPanelEditorCommands {
	/** `(addSlide)` and `(addSlideAfter)`: insert a blank slide after `index`. */
	addSlideAfter(index: number): void;
	/** `(moveSlide)`: reorder one slide. */
	moveSlide(move: SlidesPanelMoveSlide): void;
	/** `(duplicateSlides)`: duplicate every listed slide. */
	duplicateSlides(indexes: readonly number[]): void;
	/** `(deleteSlides)`: delete every listed slide (the deck keeps at least one). */
	deleteSlides(indexes: readonly number[]): void;
	/** `(toggleHideSlides)`: flip the hidden flag of every listed slide. */
	toggleHideSlides(indexes: readonly number[]): void;
	/** `(addSectionAt)`: start a new section after slide `index`. */
	addSectionAt(index: number): void;
	/** `(renameSection)`. */
	renameSection(rename: SlidesPanelRenameSection): void;
	/** `(deleteSection)`. */
	deleteSection(sectionId: string): void;
	/** `(moveSectionUp)`. */
	moveSectionUp(sectionId: string): void;
	/** `(moveSectionDown)`. */
	moveSectionDown(sectionId: string): void;
	/** `(toggleSectionCollapse)`. */
	toggleSectionCollapse(sectionId: string): void;
}

/**
 * Rail command handlers over `editor`. Each runs only while `canEdit()` is
 * true, so a host can pass its effective edit permission (for a collaboration
 * shell, `shell.canEdit`) and a stale event never edits a read-only deck.
 */
export function createSlidesPanelEditorCommands(
	editor: EditorStateService,
	canEdit: () => boolean = () => true,
): SlidesPanelEditorCommands {
	const bulk = createSlidesPaneBulkSlideOps({
		duplicateSlide: (index) => editor.duplicateSlide(index),
		deleteSlide: (index) => editor.deleteSlide(index),
		toggleSlideHidden: (index) => {
			const slide = editor.slides()[index];
			if (slide) {
				editor.updateSlide(index, { hidden: !slide.hidden });
			}
		},
	});
	const gated =
		<T extends unknown[]>(run: (...args: T) => void) =>
		(...args: T): void => {
			if (canEdit()) {
				run(...args);
			}
		};
	return {
		addSlideAfter: gated((index: number) => editor.addSlide(index)),
		moveSlide: gated((move: SlidesPanelMoveSlide) =>
			editor.moveSlide(move.fromIndex, move.toIndex),
		),
		duplicateSlides: gated((indexes: readonly number[]) => bulk.duplicateSlides(indexes)),
		deleteSlides: gated((indexes: readonly number[]) => bulk.deleteSlides(indexes)),
		toggleHideSlides: gated((indexes: readonly number[]) => bulk.toggleHideSlides(indexes)),
		addSectionAt: gated((index: number) => editor.addSection(index)),
		renameSection: gated((rename: SlidesPanelRenameSection) =>
			editor.sectionOps.rename(rename.sectionId, rename.name),
		),
		deleteSection: gated((sectionId: string) => editor.sectionOps.delete(sectionId)),
		moveSectionUp: gated((sectionId: string) => editor.sectionOps.move(sectionId, 'up')),
		moveSectionDown: gated((sectionId: string) => editor.sectionOps.move(sectionId, 'down')),
		toggleSectionCollapse: gated((sectionId: string) => editor.sectionOps.toggle(sectionId)),
	};
}
