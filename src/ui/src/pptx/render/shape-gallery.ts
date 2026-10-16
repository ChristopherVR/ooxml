/**
 * shape-gallery.ts: the Shapes gallery of Insert and Home as PowerPoint lays it out. Pure data
 * and functions: a "Recently Used Shapes" row (per viewer, kept in `localStorage`), then the nine
 * categories of the core catalogue (`ooxml-core/geometry`), with Freeform: Shape / Curve /
 * Scribble in Lines. Each item carries a thumbnail drawn from the real preset geometry, its
 * ScreenTip name key and the action a pick runs; the Insert view and the Home menu only paint it.
 *
 * @module render/shape-gallery
 */
import {
	SHAPE_GALLERY_CATEGORIES,
	shapeGalleryCategoryKey,
	shapeGalleryNameKey,
} from 'ooxml-core/geometry';
import { evaluatePresetShape } from 'ooxml-core/pptx';

import { FREEFORM_TOOL_LABEL_KEYS } from './edit-points/freeform-tool-geometry';
import type { FreeformToolKind } from './edit-points/freeform-tool-geometry';

/** What picking an item does. */
export type ShapeGalleryAction =
	| { kind: 'shape'; value: string }
	| { kind: 'actionButton'; value: string }
	| { kind: 'freeform'; value: FreeformToolKind }
	/** Listed, but the editor cannot draw it yet. */
	| { kind: 'unsupported' };

export interface ShapeGalleryItem {
	/** Stable id: the preset name, or `freeform:<tool>`. Also what the recent row stores. */
	id: string;
	labelKey: string;
	/** English name. */
	fallback: string;
	/** SVG path data on the {@link SHAPE_THUMB_VIEW_BOX} grid, outline only. */
	path: string;
	action: ShapeGalleryAction;
}

export interface ShapeGallerySection {
	id: string;
	headingKey: string;
	heading: string;
	items: readonly ShapeGalleryItem[];
}

/** The grid every thumbnail is drawn on (an 18 px box with a pixel of margin). */
export const SHAPE_THUMB_VIEW_BOX = '-1 -1 20 20';
const THUMB = 18;

export const RECENT_SHAPES_KEY = 'pptx-ui.shapes.recent';
export const RECENT_SHAPES_MAX = 8;
export const RECENT_SHAPES_HEADING_KEY = 'pptx.shapes.category.recent';

const FREEFORM_ART: Record<string, string> = {
	freeformShape: 'M2 14 L5 4 L9 10 L13 2 L16 8 L15 16 Z',
	curve: 'M2 15 C5 0 9 20 16 3',
	scribble: 'M2 12 C3 4 5 4 6 10 S9 16 10 9 S13 3 16 7',
};
const FREEFORM_FALLBACK: Record<string, string> = {
	freeformShape: 'Freeform: Shape',
	curve: 'Curve',
	scribble: 'Freeform: Scribble',
};

/** The UI token the editor inserts for a preset (a straight connector is a real connector). */
const insertToken = (preset: string): string =>
	preset === 'straightConnector1' ? 'connector' : preset;

const thumbCache = new Map<string, string>();

/** Outline path of a preset on the thumbnail grid; an empty string when it cannot be drawn. */
export function shapeThumbnailPath(preset: string): string {
	let path = thumbCache.get(preset);
	if (path === undefined) {
		path = evaluatePresetShape(preset, THUMB, THUMB)?.svgPath ?? '';
		thumbCache.set(preset, path);
	}
	return path;
}

function freeformItem(tool: 'freeformShape' | 'curve' | 'scribble'): ShapeGalleryItem {
	const known = tool !== 'scribble';
	return {
		id: `freeform:${tool}`,
		labelKey: known ? FREEFORM_TOOL_LABEL_KEYS[tool] : 'pptx.shapes.name.freeformScribble',
		fallback: FREEFORM_FALLBACK[tool] ?? tool,
		path: FREEFORM_ART[tool] ?? '',
		action: known ? { kind: 'freeform', value: tool } : { kind: 'unsupported' },
	};
}

function presetItem(preset: string, name: string): ShapeGalleryItem {
	return {
		id: preset,
		labelKey: shapeGalleryNameKey(preset),
		fallback: name,
		path: shapeThumbnailPath(preset),
		action: preset.startsWith('actionButton')
			? { kind: 'actionButton', value: preset }
			: { kind: 'shape', value: insertToken(preset) },
	};
}

const CATEGORY_SECTIONS: readonly ShapeGallerySection[] = SHAPE_GALLERY_CATEGORIES.map(
	(category) => {
		const items = category.shapes.map((shape) => presetItem(shape.preset, shape.name));
		if (category.id === 'lines') {
			// PowerPoint puts the freeform tools after the line and connector row.
			items.splice(
				4,
				0,
				freeformItem('freeformShape'),
				freeformItem('curve'),
				freeformItem('scribble'),
			);
		}
		return {
			id: category.id,
			headingKey: shapeGalleryCategoryKey(category.id),
			heading: category.name,
			items,
		};
	},
);

const ITEMS_BY_ID = new Map(
	CATEGORY_SECTIONS.flatMap((section) => section.items).map((item) => [item.id, item]),
);

/** Every item of the gallery, in category order. */
export const SHAPE_GALLERY_ITEMS: readonly ShapeGalleryItem[] = [...ITEMS_BY_ID.values()];

export function shapeGalleryItem(id: string): ShapeGalleryItem | undefined {
	return ITEMS_BY_ID.get(id);
}

/** The item a picked value stands for: an id, or the token a host got from a pick. */
export function shapeGalleryItemForValue(value: string): ShapeGalleryItem | undefined {
	return (
		ITEMS_BY_ID.get(value) ?? ITEMS_BY_ID.get(value === 'connector' ? 'straightConnector1' : '')
	);
}

/** The recently used ids, newest first; unknown or malformed storage reads as empty. */
export function readRecentShapes(): string[] {
	try {
		const parsed: unknown = JSON.parse(globalThis.localStorage?.getItem(RECENT_SHAPES_KEY) ?? '[]');
		return Array.isArray(parsed)
			? parsed
					.filter((id): id is string => typeof id === 'string' && ITEMS_BY_ID.has(id))
					.slice(0, RECENT_SHAPES_MAX)
			: [];
	} catch {
		return [];
	}
}

/** Remember a pick (an item id or the token the host received) at the front of the row. */
export function pushRecentShape(idOrValue: string): void {
	const item = shapeGalleryItemForValue(idOrValue);
	if (!item || item.action.kind === 'unsupported') {
		return;
	}
	const next = [item.id, ...readRecentShapes().filter((id) => id !== item.id)].slice(
		0,
		RECENT_SHAPES_MAX,
	);
	try {
		globalThis.localStorage?.setItem(RECENT_SHAPES_KEY, JSON.stringify(next));
	} catch {
		// Storage can be blocked or full; the row then simply stays as it was.
	}
}

/** The sections in order: Recently Used Shapes when there are any, then the categories. */
export function shapeGallerySections(
	recent: readonly string[] = readRecentShapes(),
): readonly ShapeGallerySection[] {
	const items = recent.map((id) => ITEMS_BY_ID.get(id)).filter((item) => item !== undefined);
	return items.length
		? [
				{
					id: 'recent',
					headingKey: RECENT_SHAPES_HEADING_KEY,
					heading: 'Recently Used Shapes',
					items,
				},
				...CATEGORY_SECTIONS,
			]
		: CATEGORY_SECTIONS;
}
