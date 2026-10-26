import {
	resolveThemeEffectStyle,
	resolveThemeFillStyle,
	resolveThemeLineStyle,
	type ThemeFormatScheme,
} from '../drawingml/format-scheme';
import type { DiagramDrawing } from './types';

/** Resolve inherited matrices on a render copy; authored shapes and theme references stay intact. */
export function resolveDiagramDrawingStyles(
	drawing: DiagramDrawing,
	scheme: ThemeFormatScheme | undefined,
): DiagramDrawing {
	if (!scheme) return drawing;
	return {
		...drawing,
		shapes: drawing.shapes.map((shape) => {
			const refs = shape.style;
			if (!refs) return shape;
			const fill =
				!shape.fill && refs.fill?.index !== undefined
					? refs.fill.index === 0
						? { kind: 'none' as const }
						: resolveThemeFillStyle(scheme, refs.fill.index, refs.fill.color)
					: undefined;
			const inheritedLine =
				refs.line?.index !== undefined
					? refs.line.index === 0
						? { fill: { kind: 'none' as const } }
						: resolveThemeLineStyle(scheme, refs.line.index, refs.line.color)
					: undefined;
			const line = inheritedLine ? { ...inheritedLine, ...shape.line } : undefined;
			const effects =
				!shape.effects && refs.effect?.index !== undefined
					? resolveThemeEffectStyle(scheme, refs.effect.index, refs.effect.color)
					: undefined;
			return {
				...shape,
				...(fill ? { fill } : {}),
				...(line ? { line } : {}),
				...(effects ? { effects } : {}),
			};
		}),
	};
}
