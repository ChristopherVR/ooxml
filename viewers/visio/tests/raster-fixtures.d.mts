export function rasterFixture(
	mimeType?: 'image/png' | 'image/jpeg' | 'image/gif',
	minBytes?: number,
): {
	mimeType: 'image/png' | 'image/jpeg' | 'image/gif';
	bytes: Uint8Array;
	pixelWidth: number;
	pixelHeight: number;
};
