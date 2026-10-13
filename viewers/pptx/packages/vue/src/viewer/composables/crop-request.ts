import { dispatchCropRequest } from 'ooxml-ui/pptx';

import type { PictureCropController } from './merge-crop-context';

/**
 * Run a Crop split button's request on the crop controller through the shared
 * `dispatchCropRequest`. Both crop entry points (Home > Arrange and Picture
 * Format > Size) call this.
 */
export function runCropRequest(
	controller: PictureCropController | undefined,
	value: string | number | undefined,
): void {
	if (controller) {
		dispatchCropRequest(value, {
			toggle: controller.toggleCrop,
			aspect: (width, height) => controller.applyAspect(`${width}:${height}`),
			fill: controller.applyFill,
			fit: controller.applyFit,
		});
	}
}
