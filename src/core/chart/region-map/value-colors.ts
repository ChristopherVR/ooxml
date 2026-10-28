import { parseXml } from '../../xml';
import { parseDrawingColorIn, type DrawingColor } from '../../drawingml';
import type { RegionMapOptions, RegionValueColorPosition } from './types';
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
	if (options.valueColorPositionsXml) {
		result.valueColorPositions = Array.from(root.getElementsByTagName('*'))
			.filter((e) => e.localName === 'pos')
			.flatMap((e) => {
				const kind = e.getAttribute('type');
				const value = Number(e.getAttribute('val'));
				if (!['min', 'max', 'num', 'percent'].includes(kind ?? '')) return [];
				return [
					{
						kind: kind === 'num' ? 'number' : kind,
						...(kind === 'num' || kind === 'percent' ? { value } : {}),
					} as RegionValueColorPosition,
				];
			});
	}
	return result;
}
