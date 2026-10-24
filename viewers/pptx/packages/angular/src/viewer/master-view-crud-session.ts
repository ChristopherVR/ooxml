/**
 * master-view-crud-session.ts: one Slide Master view sidebar command (Insert, Duplicate, Delete or
 * Rename a layout or master) run against the loaded deck, and its result adopted.
 *
 * `applyMasterViewCrudAction` (`ooxml-ui/pptx`) saves the deck it is given, edits the package and
 * loads it into a FRESH handler, so the result is adopted through
 * `LoadContentService.adoptMasterViewData` rather than as a plain signal patch.
 */
import { applyMasterViewCrudAction } from 'ooxml-ui/pptx';
import type {
	MasterViewCrudActionId,
	MasterViewCrudActionOptions,
	MasterViewCrudApplyResult,
	MasterViewTarget,
} from 'ooxml-ui/pptx';

import type { EditorStateService } from './editor-state.service';
import type { LoadContentService } from './load-content.service';

/** One sidebar command. */
export interface MasterViewCrudRequest {
	id: MasterViewCrudActionId;
	target: MasterViewTarget;
	options?: MasterViewCrudActionOptions;
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
	const data = loader.parsedData();
	if (!handler || !data) {
		return undefined;
	}
	const result = await applyMasterViewCrudAction(
		handler,
		data,
		request.id,
		request.target,
		request.options,
	);
	if (result.ok) {
		loader.adoptMasterViewData(result.handler, result.data);
		editor.dirty.set(true);
	}
	return result;
}
