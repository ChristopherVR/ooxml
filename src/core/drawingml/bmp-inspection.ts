import { checkImageDimensions, RasterImageError, type ImageLimits } from './raster-image';

/** Inspect a Windows BMP file before passing it to a browser image decoder. */
export function inspectBmp(view: DataView, limits: ImageLimits): [number, number] {
	const invalid = () => new RasterImageError('invalid-image', 'Invalid or truncated BMP file.');
	if (view.byteLength < 26 || view.getUint32(2, true) !== view.byteLength) throw invalid();
	const pixels = view.getUint32(10, true);
	const header = view.getUint32(14, true);
	if (header + 14 > view.byteLength || pixels < header + 14) throw invalid();
	if (![12, 40, 52, 56, 108, 124].includes(header))
		throw new RasterImageError('unsupported-image', 'Unsupported BMP header.');
	const core = header === 12;
	const width = core ? view.getUint16(18, true) : view.getInt32(18, true);
	const height = Math.abs(core ? view.getUint16(20, true) : view.getInt32(22, true));
	const depth = view.getUint16(core ? 24 : 28, true);
	const compression = core ? 0 : view.getUint32(30, true);
	if (width <= 0 || view.getUint16(core ? 22 : 26, true) !== 1) throw invalid();
	checkImageDimensions(width, height, limits);
	if (
		![1, 4, 8, 16, 24, 32].includes(depth) ||
		(compression !== 0 && !(compression === 3 && [16, 32].includes(depth) && !core))
	)
		throw new RasterImageError('unsupported-image', 'Unsupported BMP pixel format or compression.');
	const colors = depth <= 8 ? (core ? 0 : view.getUint32(46, true)) || 2 ** depth : 0;
	const maskBytes = compression === 3 && header === 40 ? 12 : 0;
	const stride = Math.ceil((width * depth) / 32) * 4;
	if (
		colors > 2 ** depth ||
		pixels < 14 + header + maskBytes + colors * (core ? 3 : 4) ||
		stride * height > view.byteLength - pixels
	)
		throw invalid();
	if (header === 124 && (view.getUint32(126, true) || view.getUint32(130, true)))
		throw new RasterImageError('unsupported-image', 'BMP color profiles are unsupported.');
	return [width, height];
}
