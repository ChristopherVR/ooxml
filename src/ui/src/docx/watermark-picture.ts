import type { PendingMediaPart } from 'ooxml-core/docx';
import { PICTURE_TYPES, pictureSize, readBytes } from './picture-commands';
import type { WatermarkPicture } from './watermark-dialog';

/** Pixels at 96 dpi to points. */
const PX_TO_PT = 0.75;

/**
 * Reads a picture chosen for a watermark and hands its bytes to `stage` under a new media part
 * name, so the next save writes them. Its natural size is in points at 96 dpi.
 */
export async function stageWatermarkPicture(
	file: File,
	partNameFor: (contentType: string) => string,
	stage: (partName: string, part: PendingMediaPart) => void,
): Promise<WatermarkPicture> {
	const contentType = file.type.toLowerCase();
	if (!PICTURE_TYPES[contentType])
		throw new Error('Choose a PNG, JPEG, GIF or BMP picture for the watermark.');
	const bytes = await readBytes(file);
	const size = await pictureSize(file);
	const partName = partNameFor(contentType);
	stage(partName, { bytes, contentType });
	const title = file.name.replace(/\.[^.]+$/, '');
	return {
		partName,
		contentType,
		widthPt: Math.max(1, size.width) * PX_TO_PT,
		heightPt: Math.max(1, size.height) * PX_TO_PT,
		...(title ? { title } : {}),
	};
}

/** The natural size in points of picture bytes already in the document, or undefined. */
export async function watermarkPictureSize(
	bytes: Uint8Array | undefined,
	contentType: string,
): Promise<{ widthPt: number; heightPt: number } | undefined> {
	if (!bytes) return undefined;
	const size = await pictureSize(new Blob([bytes as BlobPart], { type: contentType })).catch(
		() => undefined,
	);
	return size && size.width > 0
		? { widthPt: size.width * PX_TO_PT, heightPt: size.height * PX_TO_PT }
		: undefined;
}
