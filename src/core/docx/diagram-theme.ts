// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Connects the document theme to the neutral DrawingML colour resolver, so a cached SmartArt
// drawing's `a:schemeClr` colours resolve the way the rest of the document's do.
import {
	resolveDrawingColor,
	type DrawingColorTheme,
	type ResolvedDrawingColor,
} from '../drawingml/drawing-color';
import { themeColorSlotFor } from '../drawingml/theme-color';
import type { DrawingColor } from '../drawingml/types';
import type { ThemeCatalog } from './theme-model';

/** A `DrawingColorTheme` over the document theme, honouring the `clrSchemeMapping` of `settings.xml`. */
export function diagramColorTheme(theme: ThemeCatalog | undefined): DrawingColorTheme {
	return {
		scheme(name) {
			if (!theme) return undefined;
			const slot = themeColorSlotFor(name, theme.colorMapping);
			const hex = slot ? theme.colors[slot] : undefined;
			return hex ? `#${hex.replace(/^#/, '').toUpperCase()}` : undefined;
		},
	};
}

/** Resolves a colour of a diagram's cached drawing against the document theme. */
export function resolveDiagramColor(
	color: DrawingColor,
	theme: ThemeCatalog | undefined,
): ResolvedDrawingColor | undefined {
	return resolveDrawingColor(color, diagramColorTheme(theme));
}
