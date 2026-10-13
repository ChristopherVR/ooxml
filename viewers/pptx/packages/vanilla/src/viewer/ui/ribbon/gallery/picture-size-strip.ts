import type { ToolbarActionId } from 'ooxml-ui/pptx';
import { isActionHidden, pictureSizeHomeControls } from 'ooxml-ui/pptx';

import type { Translator } from '../../../i18n';
import type { CropHandlers } from '../home/crop-request';
import { runCropRequest } from '../home/crop-request';
import { createSharedHomeStrip } from '../home/shared-strip';

export interface PictureSizeState {
	editable: boolean;
	/** A single croppable picture is selected. */
	canCrop?: boolean;
	/** Picture crop mode is on. */
	cropActive?: boolean;
}

export interface PictureSizeStrip {
	el: HTMLElement;
	update(state: PictureSizeState): void;
}

/**
 * Picture Format > Size: the shared `pptx-ui-ribbon-home-picture-size` strip,
 * whose Crop split button runs the same crop handlers as Home > Arrange.
 */
export function createPictureSizeStrip(
	doc: Document,
	t: Translator,
	handlers: CropHandlers,
	hiddenActions?: readonly ToolbarActionId[],
): PictureSizeStrip {
	const strip = createSharedHomeStrip(doc, t, 'picture-size', ({ value }) =>
		runCropRequest(handlers, value),
	);
	// Presses inside the crop controls own crop mode's toggle, so they must not also commit it.
	strip.el.anchor('pictureFormat.size.crop')?.setAttribute('data-pptx-crop-controls', 'true');
	const update = (state: PictureSizeState) =>
		strip.set(
			pictureSizeHomeControls({
				editable: state.editable,
				canCrop: Boolean(state.canCrop),
				cropActive: Boolean(state.cropActive),
				hideCrop: isActionHidden('crop', hiddenActions),
			}),
		);
	update({ editable: false });
	return { el: strip.el, update };
}
