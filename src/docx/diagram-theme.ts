// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Connects the document theme to the colour resolver of the neutral `diagram` area, so a cached
// SmartArt drawing's `a:schemeClr` colours resolve the way the rest of the document's do.
import {
	resolveDrawingColor,
	type DrawingColorTheme,
	type ResolvedDrawingColor,
} from '../diagram/index.js';
import type { DiagramColor } from '../diagram/index.js';
import type { ThemeCatalog, ThemeColorSlot } from './theme-model.js';

const DEFAULT_MAPPING: Record<'bg1' | 'tx1' | 'bg2' | 'tx2', ThemeColorSlot> = {
	bg1: 'lt1',
	tx1: 'dk1',
	bg2: 'lt2',
	tx2: 'dk2',
};
const SLOTS: ReadonlySet<string> = new Set([
	'dk1',
	'lt1',
	'dk2',
	'lt2',
	'accent1',
	'accent2',
	'accent3',
	'accent4',
	'accent5',
	'accent6',
	'hlink',
	'folHlink',
]);

/** A `DrawingColorTheme` over the document theme, honouring the `clrSchemeMapping` of `settings.xml`. */
export function diagramColorTheme(theme: ThemeCatalog | undefined): DrawingColorTheme {
	return {
		scheme(name) {
			if (!theme) return undefined;
			const mapped =
				name in DEFAULT_MAPPING
					? (theme.colorMapping[name as keyof typeof DEFAULT_MAPPING] ??
						DEFAULT_MAPPING[name as keyof typeof DEFAULT_MAPPING])
					: SLOTS.has(name)
						? (name as ThemeColorSlot)
						: undefined;
			const hex = mapped ? theme.colors[mapped] : undefined;
			return hex ? `#${hex.replace(/^#/, '').toUpperCase()}` : undefined;
		},
	};
}

/** Resolves a colour of a diagram's cached drawing against the document theme. */
export function resolveDiagramColor(
	color: DiagramColor,
	theme: ThemeCatalog | undefined,
): ResolvedDrawingColor | undefined {
	return resolveDrawingColor(color, diagramColorTheme(theme));
}
