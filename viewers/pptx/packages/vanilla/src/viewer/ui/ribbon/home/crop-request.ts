import { dispatchCropRequest } from 'ooxml-ui/pptx';

/** The crop commands a Crop split button runs. */
export interface CropHandlers {
	toggleCropMode(): void;
	cropToAspect(ratioWidth: number, ratioHeight: number): void;
	cropFill(): void;
	cropFit(): void;
}

/**
 * Run a Crop split button's request on the edit handlers through the shared
 * `dispatchCropRequest`. Both crop entry points (Home > Arrange and Picture
 * Format > Size) call this.
 */
export function runCropRequest(handlers: CropHandlers, value: string | number | undefined): void {
	dispatchCropRequest(value, {
		toggle: () => handlers.toggleCropMode(),
		aspect: (width, height) => handlers.cropToAspect(width, height),
		fill: () => handlers.cropFill(),
		fit: () => handlers.cropFit(),
	});
}
