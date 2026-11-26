import type { DrawingTransform } from 'ooxml-core/docx';

export function pictureTransformCss(transform: DrawingTransform | undefined): string {
	return transform
		? `rotate(${Number(transform.rotation) || 0}deg) scale(${transform.flipH ? -1 : 1}, ${transform.flipV ? -1 : 1})`
		: '';
}
