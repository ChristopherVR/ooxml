/**
 * shape-gallery-home.ts: the Shapes gallery as rows of the Home > Drawing > Shapes menu. The same
 * sections as Insert (see `shape-gallery.ts`), with the value the host receives on a pick: a preset
 * name or the straight-connector token. Freeform tools stay on Insert, so Home lists them disabled
 * with a reason instead of faking them.
 *
 * @module render/shape-gallery-home
 */
import { shapeGallerySections } from './shape-gallery';
import type { ShapeGalleryItem } from './shape-gallery';
import type { RibbonHomeItem } from './ribbon-home-spec';

/** Ribbon-home rows for the gallery; `recent` defaults to the stored Recently Used Shapes. */
export function homeShapeGalleryItems(recent?: readonly string[]): readonly RibbonHomeItem[] {
	const rows: RibbonHomeItem[] = [];
	for (const section of shapeGallerySections(recent)) {
		for (const item of section.items) {
			rows.push(homeRow(item, section.id, section.headingKey, section.heading));
		}
	}
	return rows;
}

function homeRow(
	item: ShapeGalleryItem,
	sectionId: string,
	groupKey: string,
	group: string,
): RibbonHomeItem {
	const action = item.action;
	const value = action.kind === 'shape' || action.kind === 'actionButton' ? action.value : item.id;
	const needsInsert = action.kind === 'freeform';
	return {
		// The recent row repeats items, so its values must differ from the categories'.
		value: sectionId === 'recent' ? `recent:${value}` : value,
		labelKey: item.labelKey,
		fallback: item.fallback,
		groupKey,
		group,
		thumb: item.path,
		...(action.kind === 'shape' || action.kind === 'actionButton'
			? {}
			: {
					disabled: true,
					descriptionKey: needsInsert
						? 'pptx.shapes.freeformFromInsert'
						: 'pptx.insert.unavailable',
					description: needsInsert
						? 'Draw freeforms from Insert > Shapes.'
						: 'Not available in this viewer yet.',
				}),
	};
}

/** The preset a Home pick stands for: the `recent:` prefix of the Recently Used row is dropped. */
export const homeShapeValue = (value: string): string => value.replace(/^recent:/u, '');
