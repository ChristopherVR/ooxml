import { dispatchCropRequest } from 'ooxml-ui/pptx';

import type { PictureCropController } from '../../hooks/usePictureCropMode';

/**
 * Run a Crop split button's request on the crop controller through the shared
 * `dispatchCropRequest`. Both crop entry points (Home > Arrange and Picture
 * Format > Size) call this.
 */
export function runCropRequest(
	crop: PictureCropController | undefined,
	value: string | number | undefined,
): void {
	if (crop) {
		dispatchCropRequest(value, {
			toggle: crop.toggle,
			aspect: crop.cropToAspect,
			fill: crop.fill,
			fit: crop.fit,
		});
	}
}
