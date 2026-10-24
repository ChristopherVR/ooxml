/**
 * master-view-crud-history.ts: the undo and redo steps a Slide Master view
 * command leaves behind when it removed parts from the package.
 *
 * Delete Layout and Delete Slide Master remove the layout (or the master and
 * all its layouts) from the package, which the undo stack cannot put back.
 * The delete is allowed because no slide uses the layout NOW, but an older
 * undo step (or an undone redo step) can still have a slide on it: restoring
 * that step makes the next save write a slide relationship to a part that is
 * not in the package, and a later Insert Layout that reuses the file name
 * silently gives the slide a different layout. So after a command the binding
 * asks {@link removedSlideMasterParts} what it removed and, when anything was,
 * forgets the steps {@link historyStepUsesParts} matches
 * (`EditorHistoryController.forget`).
 *
 * @module render/master-view-crud-history
 */
import type { PptxData, PptxSlide, PptxSlideMaster } from 'ooxml-core/pptx';

/** The part of an undo or redo step that can refer to a slide master or layout. */
export interface SlideMasterHistoryStep {
	readonly slides: readonly Pick<PptxSlide, 'layoutPath'>[];
	/** The masters, in a binding whose steps restore them too. */
	readonly slideMasters?: readonly PptxSlideMaster[] | undefined;
}

function partPaths(masters: readonly PptxSlideMaster[] | undefined): string[] {
	return (masters ?? []).flatMap((master) => [
		master.path,
		...(master.layouts ?? []).map((layout) => layout.path),
	]);
}

/**
 * The slide master and layout part paths `before` has and `after` no longer
 * has: the layout Delete Layout removed, or the master Delete Slide Master
 * removed with every one of its layouts. Empty after any other command.
 */
export function removedSlideMasterParts(
	before: Pick<PptxData, 'slideMasters'>,
	after: Pick<PptxData, 'slideMasters'>,
): string[] {
	const kept = new Set(partPaths(after.slideMasters));
	return partPaths(before.slideMasters).filter((path) => !kept.has(path));
}

/**
 * True when restoring `step` would bring back a part in `removed`: one of its
 * slides is on a removed layout, or (in a step that restores the masters) its
 * masters still hold a removed master or layout, which would offer it in the
 * sidebar and the layout galleries again.
 */
export function historyStepUsesParts(
	step: SlideMasterHistoryStep,
	removed: readonly string[],
): boolean {
	if (removed.length === 0) {
		return false;
	}
	const gone = new Set(removed);
	return (
		step.slides.some((slide) => slide.layoutPath !== undefined && gone.has(slide.layoutPath)) ||
		partPaths(step.slideMasters).some((path) => gone.has(path))
	);
}
