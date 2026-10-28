/** An axis-aligned rectangle in document units. */
export interface CropRect {
	x: number;
	y: number;
	width: number;
	height: number;
}

/** The four `a:srcRect` insets as fractions of the image (negative pads). */
export interface CropInsets {
	cropLeft: number;
	cropTop: number;
	cropRight: number;
	cropBottom: number;
}

/** What a crop edit writes back onto the picture element. */
export type CropElementUpdate = CropInsets & CropRect;

/** The whole image's rectangle, derived from the frame and the insets. */
export function imageRectFromCrop(frame: CropRect, insets: CropInsets): CropRect {
	const spanX = 1 - insets.cropLeft - insets.cropRight;
	const spanY = 1 - insets.cropTop - insets.cropBottom;
	const width = spanX > 1e-6 ? frame.width / spanX : frame.width;
	const height = spanY > 1e-6 ? frame.height / spanY : frame.height;
	return {
		x: frame.x - insets.cropLeft * width,
		y: frame.y - insets.cropTop * height,
		width,
		height,
	};
}

function round4(n: number): number {
	return Math.round(n * 10000) / 10000;
}

/** The insets that show `frame` out of `image`, plus the frame itself. */
export function cropUpdateFromRects(frame: CropRect, image: CropRect): CropElementUpdate {
	return {
		x: frame.x,
		y: frame.y,
		width: frame.width,
		height: frame.height,
		cropLeft: round4((frame.x - image.x) / image.width),
		cropTop: round4((frame.y - image.y) / image.height),
		cropRight: round4((image.x + image.width - frame.x - frame.width) / image.width),
		cropBottom: round4((image.y + image.height - frame.y - frame.height) / image.height),
	};
}
