import type { VisioEdit } from '../edit-commands';
import type { VisioPage } from '../model';
import {
	VISIO_DATA_GRAPHIC_FILL_ROW,
	visioShapeDataNumber,
	visioShapeDataRow,
	type VisioDataGraphicKind,
} from './data-graphic-rules';
import {
	visioDataGraphicConfigs,
	VISIO_DATA_GRAPHIC_DEFAULT_POSITION,
	type VisioDataGraphicPosition,
} from './data-graphic-config';
import { visioDataGraphicEdits, visioRemoveDataGraphicEdits } from './data-graphics';
import {
	visioNativeGraphicAccepts,
	visioNativeGraphicItems,
	visioSetNativeGraphicEdit,
} from './data-graphic-native';
import { visioNextShapeId } from './shape-id';

/** One entry of a shape's data graphic, as the Edit Data Graphic dialog lists it. */
export interface VisioDataGraphicItem {
	kind: VisioDataGraphicKind;
	field: string;
	position?: VisioDataGraphicPosition;
	min?: number;
	max?: number;
}

/** The items of one shape: its text callouts, data bars and icon sets, then Color by Value. */
export function visioDataGraphicItems(page: VisioPage, ownerId: string): VisioDataGraphicItem[] {
	const items: VisioDataGraphicItem[] = visioDataGraphicConfigs(page, ownerId).map((config) => ({
		...config,
	}));
	const fill = visioShapeDataRow(
		page.shapes.find((shape) => shape.id === ownerId) ?? ({} as never),
		VISIO_DATA_GRAPHIC_FILL_ROW,
	);
	if (fill?.label) items.push({ kind: 'color', field: fill.label });
	return items;
}

const nextAfter = (edits: readonly VisioEdit[], start: number) =>
	edits.reduce((next, edit) => {
		const id = 'shapeId' in edit ? Number(edit.shapeId) : NaN;
		return Number.isFinite(id) ? Math.max(next, id + 1) : next;
	}, start);

/**
 * Edits that add an item to the graphics of these shapes, or change `replacing` into it. An item
 * is a kind and a field: a shape can show several text callouts or bars for different fields.
 * Items that share a place stack, and a bar or icon range defaults to every selected value.
 */
export function visioSetDataGraphicItemEdits(
	page: VisioPage,
	ownerIds: readonly string[],
	item: VisioDataGraphicItem,
	replacing?: Pick<VisioDataGraphicItem, 'kind' | 'field'>,
): VisioEdit[] {
	const owners = page.shapes.filter((shape) => ownerIds.includes(shape.id));
	const values = owners
		.map((shape) => visioShapeDataNumber(shape, item.field))
		.filter((value): value is number => value !== undefined);
	const min = item.min ?? Math.min(0, ...values);
	const max = item.max ?? Math.max(...values, min + 1);
	const edits: VisioEdit[] = [];
	let next = Number(visioNextShapeId(page));
	for (const owner of owners) {
		if (item.kind !== 'color' && visioNativeGraphicAccepts(page, owner)) {
			const kept = visioNativeGraphicItems(page, owner.id).filter(
				(config) =>
					!(config.kind === item.kind && config.field === item.field) &&
					!(config.kind === replacing?.kind && config.field === replacing?.field),
			);
			edits.push(
				...(replacing?.kind === 'color'
					? visioRemoveDataGraphicEdits(page, [owner.id], 'color', replacing.field)
					: []),
				visioSetNativeGraphicEdit(page, owner.id, [
					...kept,
					{
						kind: item.kind,
						field: item.field,
						...(item.position ? { position: item.position } : {}),
						...(item.kind === 'text' ? {} : { min, max }),
					},
				]),
			);
			continue;
		}
		const others = visioDataGraphicConfigs(page, owner.id).filter(
			(config) =>
				!(config.kind === item.kind && config.field === item.field) &&
				!(config.kind === replacing?.kind && config.field === replacing?.field),
		);
		const at = item.position ?? VISIO_DATA_GRAPHIC_DEFAULT_POSITION[item.kind as 'text'];
		const slot = others.filter(
			(config) => (config.position ?? VISIO_DATA_GRAPHIC_DEFAULT_POSITION[config.kind]) === at,
		).length;
		const removed =
			item.kind === 'color'
				? []
				: [
						...visioRemoveDataGraphicEdits(page, [owner.id], item.kind, item.field),
						...(replacing && (replacing.kind !== item.kind || replacing.field !== item.field)
							? visioRemoveDataGraphicEdits(page, [owner.id], replacing.kind, replacing.field)
							: []),
					];
		const made = visioDataGraphicEdits(page, [owner.id], {
			kind: item.kind,
			field: item.field,
			...(item.position ? { position: item.position } : {}),
			...(item.kind === 'text' || item.kind === 'color' ? {} : { min, max }),
			...(slot ? { slot } : {}),
			idStart: next,
			replace: 'none',
		});
		const fill = visioShapeDataRow(owner, VISIO_DATA_GRAPHIC_FILL_ROW);
		// Colouring by another field keeps the fill remembered from before any colouring.
		const recolor =
			item.kind === 'color' && fill && fill.label !== item.field
				? [
						{
							type: 'set-shape-data',
							pageId: page.id,
							shapeId: owner.id,
							row: fill.name,
							data: {
								label: item.field,
								prompt: '',
								type: 'string',
								format: '',
								value: String(fill.value ?? ''),
								invisible: true,
							},
						} as VisioEdit,
					]
				: [];
		edits.push(...removed, ...made, ...recolor);
		next = nextAfter(made, next);
	}
	return edits;
}

/** Edits that remove one item from the graphics of these shapes. */
export const visioDeleteDataGraphicItemEdits = (
	page: VisioPage,
	ownerIds: readonly string[],
	item: Pick<VisioDataGraphicItem, 'kind' | 'field'>,
): VisioEdit[] => visioRemoveDataGraphicEdits(page, ownerIds, item.kind, item.field);
