// Writes the shapes of one page: a shell entry per shape (nested shapes name their parent) and an
// entry per cell, section, row and other child, so edits to different properties of one shape are
// different entries.
import * as Y from 'yjs';
import { cellId, entryContent, shapeId, shellContent } from './baseline';
import type { DrawingKeys } from './keys';
import { assignPositions } from './order';
import type { PieceEntry } from './pieces';
import {
	type CellEntry,
	type ShapeEntry,
	type SharedDrawing,
	cellEntry,
	shapeEntry,
} from './schema';
import type { PageSnapshot } from './snapshot';

export interface WriteContext {
	shared: SharedDrawing;
	keys: DrawingKeys;
	/** Baseline of what is written or kept by this pass. */
	now: Map<string, string>;
	/** Shared by every entry created or rewritten in this pass. */
	batch: string;
	write(map: Y.Map<unknown>, key: string, value: unknown): void;
}

/**
 * Writes `local()` when it differs from this peer's baseline (a local edit) or the room has no
 * entry; otherwise keeps the room's value (a collaborator's change not applied here yet), with the
 * local position. `stamp` records the batch on an entry that is written.
 */
export function syncEntry<T extends { pos?: number; b?: string }>(
	ctx: WriteContext,
	map: Y.Map<unknown>,
	mapKey: string,
	id: string,
	content: string,
	room: T | undefined,
	roomContent: () => string,
	local: () => T,
	stamp = false,
): void {
	ctx.now.set(id, content);
	const base = ctx.keys.known.get(id);
	const own = base === undefined ? !room || roomContent() !== content : base !== content;
	if (own || !room) {
		const value = local();
		ctx.write(map, mapKey, stamp ? { ...value, b: ctx.batch } : value);
	} else {
		const pos = local().pos;
		ctx.write(map, mapKey, pos === undefined ? room : { ...room, pos });
	}
}

/** Positions for `keys` grouped by `group`, keeping each group's existing increasing order. */
function groupPositions(
	items: readonly { key: string; group: string }[],
	current: (key: string) => number | undefined,
): Map<string, number> {
	const groups = new Map<string, string[]>();
	for (const item of items) groups.set(item.group, [...(groups.get(item.group) ?? []), item.key]);
	const out = new Map<string, number>();
	for (const list of groups.values())
		for (const [key, pos] of assignPositions(list, current)) out.set(key, pos);
	return out;
}

export function syncShapes(ctx: WriteContext, page: PageSnapshot): void {
	const { shared } = ctx;
	const shapePos = groupPositions(
		page.shapes.map((s) => ({ key: s.key, group: s.parent ?? '' })),
		(key) => shapeEntry(shared.shapes.get(`${page.key}/${key}`))?.pos,
	);
	for (const shape of page.shapes) {
		const full = `${page.key}/${shape.key}`;
		const room = shapeEntry(shared.shapes.get(full));
		const parent = shape.parent;
		syncEntry<ShapeEntry>(
			ctx,
			shared.shapes,
			full,
			shapeId(page.key, shape.key),
			shellContent({ parent, open: shape.open }),
			room,
			() => shellContent(room!),
			() => ({
				pos: shapePos.get(shape.key)!,
				...(parent !== undefined ? { parent } : {}),
				open: shape.open,
				b: room?.b ?? ctx.batch,
			}),
		);
		const piecePos = groupPositions(
			shape.pieces.map((p) => ({ key: p.path, group: p.parent })),
			(path) => cellEntry(shared.cells.get(`${full}/${path}`))?.pos,
		);
		for (const piece of shape.pieces) {
			const cellKey = `${full}/${piece.path}`;
			const roomCell = cellEntry(shared.cells.get(cellKey));
			const local = (): PieceEntry => ({
				pos: piecePos.get(piece.path)!,
				...(piece.kind === 'leaf' ? { xml: piece.text } : { open: piece.text }),
			});
			syncEntry<CellEntry>(
				ctx,
				shared.cells,
				cellKey,
				cellId(page.key, shape.key, piece.path),
				entryContent(piece.kind === 'leaf' ? { xml: piece.text } : { open: piece.text }),
				roomCell,
				() => entryContent(roomCell!),
				local,
			);
		}
	}
}
