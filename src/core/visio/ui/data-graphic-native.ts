import type { VisioEdit } from '../edit-commands';
import type { VisioPage, VisioShape } from '../model';
import type { VisioDataGraphicItemSpec } from '../edit-data-graphic';
import { visioDataGraphicPositionOf, type VisioDataGraphicPosition } from '../data-graphic-place';
import type { VisioDataGraphicConfig } from './data-graphic-config';
import { visioDataGraphicMarker, visioDataGraphicParts } from './data-graphic-rules';

/** The User row Visio keeps on a shape that has a data graphic: it names the graphic's master. */
export const VISIO_DG_DISPLAY_ROW = 'visDGDisplayFormat';

const user = (shape: VisioShape, name: string): string | undefined =>
	shape.userCells?.find((cell) => cell.name === name)?.value;

/** True for a shape that carries a data graphic made of callout instances. */
export const visioHasNativeGraphic = (shape: VisioShape): boolean =>
	user(shape, VISIO_DG_DISPLAY_ROW) !== undefined && shape.children.some(isCallout);

const isCallout = (shape: VisioShape): boolean => user(shape, 'visDGField0') !== undefined;

const KINDS: Record<string, VisioDataGraphicConfig['kind']> = {
	'Text Callout': 'text',
	'Data Bar': 'bar',
	'Icon Set': 'icon',
};

/** The callouts of one shape's data graphic, in the order they are stacked. */
export function visioNativeGraphicItems(
	page: VisioPage,
	ownerId: string,
): VisioDataGraphicConfig[] {
	const owner = page.shapes.find((shape) => shape.id === ownerId);
	if (!owner || !visioHasNativeGraphic(owner)) return [];
	return owner.children.filter(isCallout).flatMap((child) => {
		const kind = KINDS[user(child, 'msvCalloutType') ?? ''];
		const field = /^\{(.*)\}$/.exec(user(child, 'visDGField0') ?? '')?.[1];
		if (!kind || !field) return [];
		const point = /^PNT\((-?\d+),(-?\d+)\)$/.exec(user(child, 'visDGDefaultPos') ?? '');
		const position = point
			? visioDataGraphicPositionOf(Number(point[1]), Number(point[2]))
			: undefined;
		const range = (name: string) => {
			const value = child.shapeData?.find((row) => row.name === name)?.value;
			return typeof value === 'number' ? value : undefined;
		};
		const [min, max] = [range('msvCalloutPropMin'), range('msvCalloutPropMax')];
		return [
			{
				kind,
				field,
				...(position ? { position } : {}),
				...(kind === 'text' || min === undefined ? {} : { min }),
				...(kind === 'text' || max === undefined ? {} : { max }),
			},
		];
	});
}

/** Shapes whose graphic is a native one. */
export const visioNativeGraphicOwners = (page: VisioPage): string[] =>
	page.shapes.filter(visioHasNativeGraphic).map((shape) => shape.id);

/**
 * Whether the next text, bar or icon item on this shape is written natively: a local shape (or
 * the group a native graphic made) that is not already drawn with marked shapes. Instances of a
 * stencil master and ordinary groups keep the marked-shape drawing.
 */
export function visioNativeGraphicAccepts(page: VisioPage, shape: VisioShape): boolean {
	if (visioDataGraphicMarker(shape) || shape.masterId) return false;
	if (visioHasNativeGraphic(shape)) return true;
	if (visioDataGraphicParts(page).get(shape.id)?.length) return false;
	return shape.kind === 'shape' && !shape.children.length && shape.geometry.length > 0;
}

const nativeSpec = (config: VisioDataGraphicConfig): VisioDataGraphicItemSpec => ({
	kind: config.kind,
	field: config.field,
	...(config.position ? { position: config.position as VisioDataGraphicPosition } : {}),
	...(config.min === undefined ? {} : { min: config.min }),
	...(config.max === undefined ? {} : { max: config.max }),
});

/** The edit that makes `items` the shape's whole data graphic. */
export const visioSetNativeGraphicEdit = (
	page: VisioPage,
	ownerId: string,
	items: readonly VisioDataGraphicConfig[],
): VisioEdit => ({
	type: 'set-data-graphic',
	pageId: page.id,
	shapeId: ownerId,
	items: items.map(nativeSpec),
});
