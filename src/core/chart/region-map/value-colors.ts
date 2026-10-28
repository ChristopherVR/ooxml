import { parseXml } from '../../xml';
import { parseDrawingColorIn, type DrawingColor } from '../../drawingml';
import type { RegionMapOptions } from './types';
import { parseRegionValueColorPositions } from './value-color-positions';
export function resolveRegionValueColors(
	options: RegionMapOptions,
	color: (c: DrawingColor | undefined) => string | undefined,
): RegionMapOptions {
	const result = { ...options };
	const xml = `<cx:root xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${options.valueColorsXml ?? ''}${options.valueColorPositionsXml ?? ''}</cx:root>`;
	const root = parseXml(xml).documentElement;
	if (options.valueColorsXml) {
		const colors = Array.from(root.getElementsByTagName('*')).flatMap((node) => {
			const c = color(parseDrawingColorIn(node));
			return c ? [c] : [];
		});
		if (colors.length) result.valueColors = colors;
	}
	const positions = parseRegionValueColorPositions(options.valueColorPositionsXml);
	if (positions) result.valueColorPositions = positions;
	return result;
}
