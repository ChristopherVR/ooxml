import { evaluatePresetShape } from '../../geometry/preset-shape-evaluator';
import type { ImageObject } from '../model';

/** The picture's preset outline in grid pixels. Rectangles use the normal box painter. */
export function pictureFramePath(
	picture: ImageObject,
	width: number,
	height: number,
): string | undefined {
	if (!picture.geometry || picture.geometry === 'rect') return undefined;
	return (
		evaluatePresetShape(picture.geometry, width, height, picture.adjustments)
			?.paths.map(({ d }) => d)
			.join(' ') || undefined
	);
}
