import type { ResolvedDrawingColor } from './drawing-color';

/** CSS/SVG paint retains the alpha resolved by the shared DrawingML color engine. */
export function drawingColorCss(value: ResolvedDrawingColor | undefined): string | undefined {
	if (!value || value.alpha === 1) return value?.hex;
	const channels = [1, 3, 5].map((offset) =>
		Number.parseInt(value.hex.slice(offset, offset + 2), 16),
	);
	return `rgba(${channels.join(',')},${value.alpha})`;
}
