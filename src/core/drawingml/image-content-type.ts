export function imageContentType(partName: string): string {
	const extension = partName.split('.').pop()?.toLowerCase();
	const known: Record<string, string> = {
		png: 'image/png',
		jpg: 'image/jpeg',
		jpeg: 'image/jpeg',
		gif: 'image/gif',
		bmp: 'image/bmp',
		svg: 'image/svg+xml',
		tif: 'image/tiff',
		tiff: 'image/tiff',
		emf: 'image/x-emf',
		wmf: 'image/x-wmf',
	};
	return (extension && known[extension]) || 'application/octet-stream';
}
