export const ZOOM_COVER_TYPES = 'image/png,image/jpeg,image/webp,image/gif';

export async function zoomCover(doc: Document, file: File): Promise<string> {
	if (file.size > 16 * 1024 * 1024) throw new Error('The Zoom cover image is too large.');
	const bitmap = await createImageBitmap(file);
	try {
		const scale = Math.min(1, 360 / Math.max(bitmap.width, bitmap.height));
		const canvas = doc.createElement('canvas');
		canvas.width = Math.max(1, Math.round(bitmap.width * scale));
		canvas.height = Math.max(1, Math.round(bitmap.height * scale));
		const context = canvas.getContext('2d');
		if (!context) throw new Error('Could not render the Zoom cover image.');
		context.drawImage(bitmap, 0, 0, canvas.width, canvas.height);
		return canvas.toDataURL('image/png');
	} finally {
		bitmap.close();
	}
}
