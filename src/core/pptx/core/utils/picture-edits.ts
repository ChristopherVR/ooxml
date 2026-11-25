import type {
	PptxBackgroundRemoval,
	PptxElement,
	PptxImageLikeElement,
	ShapeStyle,
} from '../types';

/** Replace only the raster and its SVG/source associations. */
export function replacePicturePatch(imageData: string): Partial<PptxImageLikeElement> {
	return {
		imageData,
		imagePath: undefined,
		imageDataEdited: true,
		svgData: undefined,
		svgPath: undefined,
		originalPicturePath: undefined,
		originalPictureData: undefined,
	};
}

export function removePictureBackgroundPatch(
	picture: PptxImageLikeElement,
	imageData: string,
	originalPictureData: string,
	removal: PptxBackgroundRemoval,
): Partial<PptxImageLikeElement> {
	return {
		...replacePicturePatch(imageData),
		originalPictureData,
		imageEffects: { ...picture.imageEffects, backgroundRemoval: removal },
	};
}

export function clearPictureCrop(): Partial<PptxImageLikeElement> {
	return { cropLeft: 0, cropRight: 0, cropTop: 0, cropBottom: 0 };
}

/** Retain the position and frame unless the caller supplies the source pixel dimensions. */
export function resetPicturePatch(
	picture: PptxImageLikeElement,
	options: { original?: string; size?: { width: number; height: number } } = {},
): Partial<PptxImageLikeElement> {
	return {
		...(options.original ? replacePicturePatch(options.original) : {}),
		...options.size,
		...clearPictureCrop(),
		imageEffects: undefined,
		shapeStyle: {
			styleMatrixReset: true,
			effectListXml: {},
			fillMode: 'none',
			strokeFillMode: 'none',
			strokeWidth: 0,
			strokeColor: 'transparent',
		},
		...(picture.cropShape && picture.cropShape !== 'none'
			? { cropShape: 'none', shapeType: 'rect', shapeAdjustments: undefined }
			: {}),
	};
}

export function pictureBorderPatch(
	picture: PptxImageLikeElement,
	border: { color: string; width: number; dash: NonNullable<ShapeStyle['strokeDash']> },
): Partial<PptxImageLikeElement> {
	const shapeStyle: ShapeStyle = { ...picture.shapeStyle };
	delete shapeStyle.strokeColorXml;
	delete shapeStyle.strokeColorRef;
	delete shapeStyle.strokeOpacity;
	delete shapeStyle.strokeGradientXml;
	delete shapeStyle.strokeGradientStops;
	delete shapeStyle.strokePatternXml;
	delete shapeStyle.customDashSegments;
	delete shapeStyle.customDashSegmentXml;
	delete shapeStyle.customDashXml;
	return {
		shapeStyle: {
			...shapeStyle,
			strokeColor: border.color,
			strokeWidth: border.width,
			strokeDash: border.dash,
			strokeFillMode: border.width > 0 ? 'solid' : 'none',
		},
	};
}

export function isDirectPicture(
	element: PptxElement | null | undefined,
): element is PptxImageLikeElement {
	return element?.type === 'picture' || element?.type === 'image';
}
