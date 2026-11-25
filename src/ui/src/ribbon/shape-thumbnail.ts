import { evaluatePresetShape } from 'ooxml-core/geometry';

const paths = new Map<string, string>();
export const SHAPE_THUMB_VIEW_BOX = '-1 -1 20 20';

export function shapeThumbnailPath(preset: string): string {
	const cached = paths.get(preset);
	if (cached !== undefined) return cached;
	const path = evaluatePresetShape(preset, 18, 18)?.svgPath ?? '';
	paths.set(preset, path);
	return path;
}
