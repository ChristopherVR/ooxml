import { cropFill, cropFit, cropToAspectRatio, dispatchCropRequest } from 'ooxml-ui/pptx';
import type { CropElementUpdate, NaturalImageSize } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';

import type { EditorState } from '../../../editor/editor-state.svelte';

/** The rendered bitmap's natural size, when its <img> is on the canvas. */
function naturalSizeOf(element: PptxElement): NaturalImageSize | undefined {
	const img =
		typeof document === 'undefined'
			? null
			: document.querySelector<HTMLImageElement>(
					`[data-element-id="${CSS.escape(element.id)}"] img`,
				);
	return img && img.naturalWidth > 0 && img.naturalHeight > 0
		? { width: img.naturalWidth, height: img.naturalHeight }
		: undefined;
}

/**
 * Run a Crop split button's request through the shared `dispatchCropRequest`:
 * the main half toggles crop mode, a menu pick applies its aspect ratio, Fill
 * or Fit to the selected picture as one undoable update. Both crop entry
 * points (Home > Arrange and Picture Format > Size) call this.
 */
export function runCropRequest(editor: EditorState, value: string | number | undefined): void {
	const apply = (compute: (element: PptxElement) => CropElementUpdate) => {
		const element = editor.selectedElement;
		if (element) {
			editor.cropOps.applyOnce(compute(element));
		}
	};
	dispatchCropRequest(value, {
		toggle: () => editor.cropOps.toggle(),
		aspect: (width, height) => apply((el) => cropToAspectRatio(el, width, height)),
		fill: () => apply((el) => cropFill(el, naturalSizeOf(el))),
		fit: () => apply((el) => cropFit(el, naturalSizeOf(el))),
	});
}
