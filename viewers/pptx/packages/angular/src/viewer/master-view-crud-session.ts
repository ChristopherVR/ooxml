/**
 * master-view-crud-session.ts: one Slide Master view sidebar command (Insert, Duplicate, Delete or
 * Rename a layout or master) run on the deck as it is now, and its result adopted in place, by the
 * rule `applyMasterViewCrudAction`'s doc (`ooxml-ui/pptx`) gives every binding.
 *
 * What is Angular's own: the slides it hands the command come from the editor (the loader's parsed
 * deck is the deck as loaded), and `loader.slides` is left alone because the viewer re-seeds
 * itself from it (undo history cleared, the deck marked saved, back to the first slide).
 */
import { applyMasterViewCrudAction } from 'ooxml-ui/pptx';
import type {
	MasterViewCrudActionId,
	MasterViewCrudActionOptions,
	MasterViewCrudApplyResult,
	MasterViewTarget,
} from 'ooxml-ui/pptx';
import type { PptxData, PptxSlide } from 'pptx-viewer-core';

import type { EditorStateService } from './editor-state.service';
import type { LoadContentService } from './load-content.service';

/** One sidebar command. */
export interface MasterViewCrudRequest {
	id: MasterViewCrudActionId;
	target: MasterViewTarget;
	options?: MasterViewCrudActionOptions;
	/** The deck's slides as a save would write them now (template elements merged back). */
	slides: readonly PptxSlide[];
}

/**
 * The deck the sidebar commands read: `slides` and the masters as they are now, over the parsed
 * deck. The command list asks it which layouts a slide uses, so a layout an edit started using
 * cannot be deleted.
 */
export function masterViewDeck(loader: LoadContentService, slides: readonly PptxSlide[]): PptxData {
	return {
		...(loader.parsedData() ?? { width: 0, height: 0 }),
		slides: [...slides],
		slideMasters: loader.slideMasters(),
	};
}

/**
 * Run `request` and adopt its result; `undefined` when no deck is loaded. A failed command
 * changes nothing and hands back the failure for the caller to report.
 */
export async function runMasterViewCrud(
	loader: LoadContentService,
	editor: EditorStateService,
	request: MasterViewCrudRequest,
): Promise<MasterViewCrudApplyResult | undefined> {
	const handler = loader.getHandler();
	if (!handler || !loader.parsedData()) {
		return undefined;
	}
	const result = await applyMasterViewCrudAction(
		handler,
		masterViewDeck(loader, request.slides),
		request.id,
		request.target,
		request.options,
	);
	if (result.ok) {
		loader.adoptMasterViewData(result.handler, result.data, true);
		editor.dirty.set(true);
	}
	return result;
}
