import type { ToolbarActionId } from 'ooxml-ui/pptx';

import type { EditActions } from '../../../editor/editor-edit-ops';
import type { Translator } from '../../../i18n';
import { createArrangeGroup } from './arrange-group';

export function createContextualArrange(
	doc: Document,
	t: Translator,
	edit: EditActions,
	hiddenActions?: readonly ToolbarActionId[],
) {
	return createArrangeGroup(
		doc,
		t,
		{
			bringForward: edit.bringForward,
			sendBackward: edit.sendBackward,
			bringToFront: edit.bringToFront,
			sendToBack: edit.sendToBack,
			alignElements: edit.alignElements,
			distributeElements: edit.distributeElements,
			flipHorizontal: edit.flipHorizontal,
			flipVertical: edit.flipVertical,
			groupSelected: edit.groupSelected,
			ungroupSelected: edit.ungroupSelected,
			setStrokeWidth: edit.setShapeStrokeWidth,
			toggleFormatPainter: edit.toggleFormatPainter,
			duplicate: edit.duplicateSelected,
			delete: edit.deleteSelected,
			mergeShapes: edit.mergeShapes,
			toggleCropMode: edit.toggleCropMode,
			cropToAspect: edit.cropToAspect,
			cropFill: edit.cropFill,
			cropFit: edit.cropFit,
		},
		hiddenActions,
		true,
	);
}
