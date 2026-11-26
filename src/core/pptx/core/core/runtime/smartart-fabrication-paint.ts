import type { DrawingColor, DrawingFill, DrawingLine } from '../../../../drawingml/types';
import type { PptxSmartArtDrawingShape } from '../../types';

function color(hex: string | undefined, opacity = 1): DrawingColor | undefined {
	const value = hex?.replace(/^#/u, '').trim();
	if (!value || !/^[0-9a-f]{6}$/iu.test(value)) return;
	return {
		kind: 'srgb',
		value: value.toUpperCase(),
		transforms:
			opacity < 1
				? [{ name: 'alpha', value: String(Math.round(Math.max(0, opacity) * 100000)) }]
				: [],
	};
}

export function smartArtDrawingFill(shape: PptxSmartArtDrawingShape): DrawingFill | undefined {
	if (shape.pictureFill) return shape.pictureFill;
	if (shape.fillNone) return { kind: 'none' };
	const stops = shape.fillGradientStops?.flatMap((stop) => {
		const stopColor = color(stop.color, stop.opacity);
		return stopColor ? [{ position: stop.position, color: stopColor }] : [];
	});
	if (stops && stops.length > 1) {
		return {
			kind: 'gradient',
			stops,
			...(shape.fillGradientType === 'radial'
				? { path: 'circle' }
				: { angle: shape.fillGradientAngle ?? 0 }),
		};
	}
	const solid = color(shape.fillColor, shape.fillOpacity);
	return solid ? { kind: 'solid', color: solid } : undefined;
}

export function smartArtDrawingLine(shape: PptxSmartArtDrawingShape): DrawingLine | undefined {
	if (shape.strokeColor === 'none') return { fill: { kind: 'none' } };
	const solid = color(shape.strokeColor);
	return (
		solid && {
			...(shape.strokeWidth && shape.strokeWidth > 0
				? { widthEmu: Math.round(shape.strokeWidth * 12700) }
				: {}),
			fill: { kind: 'solid', color: solid },
		}
	);
}
