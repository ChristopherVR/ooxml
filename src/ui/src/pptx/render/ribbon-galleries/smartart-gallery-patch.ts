/**
 * The SmartArt style / colour-scheme patch every binding's inspector and the
 * SmartArt Design galleries share: merge the change into `smartArtData`, then
 * rebuild drawing shapes a structural edit had cleared (exactly what React's
 * `SmartArtPropertiesPanel.applySmartArtData` does with the element's box).
 *
 * Also the gallery catalogues: the core model offers five colour schemes and
 * three style intensities; each is named after the PowerPoint gallery entry
 * it writes (`Application.SmartArtColors` / `SmartArtQuickStyles` names, and
 * the accents `smartart-fabrication-styles.ts` puts in the fabricated
 * `dgm:colorsDef`).
 *
 * @module render/ribbon-galleries/smartart-gallery-patch
 */
import { switchSmartArtLayout } from 'ooxml-core/pptx';
import type {
	PptxElement,
	SmartArtLayoutType,
	PptxSmartArtData,
	PptxThemeColorScheme,
	SmartArtColorScheme,
	SmartArtStyle,
} from 'ooxml-core/pptx';

import { builtinLayoutPatch, smartArtBuiltinLayouts } from '../smartart-builtin-layouts';
import { resolvePalette } from '../smartart-drawing';
import { rebuildDrawingShapesIfCleared } from '../smartart-reflow-to-shapes';

/** `data` with `patch` merged, drawing shapes rebuilt when a structural edit cleared them. */
export function applySmartArtDataPatch(
	data: PptxSmartArtData,
	patch: Partial<PptxSmartArtData>,
	box?: { width: number; height: number },
	elementId = 'inspector',
): PptxSmartArtData {
	const next = { ...data, ...patch };
	if (patch.style !== undefined) {
		next.quickStyleDirty = true;
		next.quickStyle = {
			...next.quickStyle,
			effectIntensity: patch.style === 'flat' ? 'subtle' : patch.style,
		};
	}
	return box
		? rebuildDrawingShapesIfCleared(
				next,
				next.layout,
				resolvePalette(next),
				next.style ?? 'flat',
				elementId,
				box,
			)
		: next;
}

/**
 * The fields a layout switch changes, for a binding that sends a patch rather
 * than the whole data. A switch also CLEARS the previous layout's parsed
 * `layoutDefinition` and `presLayoutVars`; a patch that leaves them out keeps
 * them on the element, and the layout engine then runs the old definition ahead
 * of the new family (the diagram keeps its old arrangement or breaks).
 * `drawingShapes` is forwarded cleared so the shapes are rebuilt.
 */
export function smartArtLayoutSwitchPatch(
	data: PptxSmartArtData,
	layout: SmartArtLayoutType,
): Partial<PptxSmartArtData> {
	const updated = switchSmartArtLayout(data, layout);
	// A bare category used to leave no layout definition, so a coarse family renderer drew the
	// diagram. Apply the family's own built-in layout so the real DiagramML engine lays it out.
	const defaultId = smartArtBuiltinLayouts()?.defaultBuiltinSmartArtLayoutId(layout);
	const builtin = defaultId ? builtinLayoutPatch(data, defaultId) : undefined;
	if (builtin) {
		return { ...builtin, resolvedLayoutType: layout, layoutType: layout };
	}
	return {
		layoutType: updated.layoutType,
		resolvedLayoutType: updated.resolvedLayoutType,
		layout: updated.layout,
		layoutDefinition: updated.layoutDefinition,
		builtinLayoutId: updated.builtinLayoutId,
		presLayoutVars: updated.presLayoutVars,
		layoutDirty: updated.layoutDirty,
		drawingDirty: updated.drawingDirty,
		drawingShapes: updated.drawingShapes,
	};
}

/**
 * `data` switched to layout `layout`: {@link smartArtLayoutSwitchPatch} merged onto it. For a
 * binding that applies whole data rather than a patch; every binding's layout switcher uses it so
 * all of them lay the diagram out with the family's built-in layout.
 */
export function switchSmartArtLayoutData(
	data: PptxSmartArtData,
	layout: SmartArtLayoutType,
): PptxSmartArtData {
	if ((data.resolvedLayoutType ?? 'list') === layout) {
		return data;
	}
	return { ...data, ...smartArtLayoutSwitchPatch(data, layout) };
}

/** The element patch for a SmartArt style / colour change, or null for a non-SmartArt. */
export function smartArtElementPatch(
	element: PptxElement | null,
	patch: Partial<PptxSmartArtData>,
): { elementId: string; patch: Partial<PptxElement> } | null {
	if (element?.type !== 'smartArt' || !element.smartArtData) {
		return null;
	}
	const smartArtData = applySmartArtDataPatch(
		element.smartArtData,
		patch,
		{ width: element.width, height: element.height },
		element.id,
	);
	return { elementId: element.id, patch: { smartArtData } as Partial<PptxElement> };
}

export interface SmartArtColorEntry {
	id: SmartArtColorScheme;
	section: 'colorful' | 'accent1' | 'accent2';
	/** PowerPoint's gallery name. */
	name: string;
	labelKey: string;
	/** Theme accents the scheme fills nodes with, in cycle order. */
	accents: ReadonlyArray<keyof PptxThemeColorScheme>;
}

export const SMARTART_COLOR_ENTRIES: readonly SmartArtColorEntry[] = [
	{
		id: 'colorful1',
		section: 'colorful',
		name: 'Colorful - Accent Colors',
		labelKey: 'pptx.gallery.smartArtColors.colorful1',
		accents: ['accent1', 'accent2', 'accent3'],
	},
	{
		id: 'colorful2',
		section: 'colorful',
		name: 'Colorful Range - Accent Colors 2 to 3',
		labelKey: 'pptx.gallery.smartArtColors.colorfulRange',
		accents: ['accent2', 'accent3', 'accent4'],
	},
	{
		id: 'colorful3',
		section: 'colorful',
		name: 'Colorful Range - Accent Colors 3 to 4',
		labelKey: 'pptx.gallery.smartArtColors.colorfulRange',
		accents: ['accent3', 'accent4', 'accent5'],
	},
	{
		id: 'monochromatic1',
		section: 'accent1',
		name: 'Colored Fill - Accent 1',
		labelKey: 'pptx.gallery.smartArtColors.coloredFill',
		accents: ['accent1'],
	},
	{
		id: 'monochromatic2',
		section: 'accent2',
		name: 'Colored Fill - Accent 2',
		labelKey: 'pptx.gallery.smartArtColors.coloredFill',
		accents: ['accent2'],
	},
];

export interface SmartArtStyleEntry {
	id: SmartArtStyle;
	name: string;
	labelKey: string;
}

export const SMARTART_STYLE_ENTRIES: readonly SmartArtStyleEntry[] = [
	{ id: 'flat', name: 'Simple Fill', labelKey: 'pptx.gallery.smartArtStyles.simpleFill' },
	{
		id: 'moderate',
		name: 'Moderate Effect',
		labelKey: 'pptx.gallery.smartArtStyles.moderateEffect',
	},
	{ id: 'intense', name: 'Intense Effect', labelKey: 'pptx.gallery.smartArtStyles.intenseEffect' },
];

/** The patch that applies built-in layout `layoutId`, or `undefined` while the library loads. */
export function smartArtBuiltinLayoutPatch(
	data: PptxSmartArtData,
	layoutId: string,
): Partial<PptxSmartArtData> | undefined {
	return builtinLayoutPatch(data, layoutId);
}
