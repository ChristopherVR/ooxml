import { editVsdxDataGraphics, type VisioSetDataGraphicEdit } from './edit-data-graphic';
import { readDataGraphicItems } from './edit-data-graphic-refresh';
import type { VisioEdit } from './edit-commands';
import { openEditablePackage } from './edit-package';
import { shapesOf } from './edit-master-add';
import type { VisioPackageLimits } from './package-common';
import { visioXml } from './parts';
import { attribute } from './sheet';

/** A shape an edit acts on, and the shapes that carry its data graphic once the edit is done. */
interface Target {
	pageId: string;
	shapeId: string;
	keep: string[];
}

function targetsOf(edits: readonly VisioEdit[]): Target[] {
	const targets: Target[] = [];
	for (const edit of edits) {
		if (edit.type === 'resize-shape')
			targets.push({ pageId: edit.pageId, shapeId: edit.shapeId, keep: [edit.shapeId] });
		else if (edit.type === 'delete-shape')
			targets.push({ pageId: edit.pageId, shapeId: edit.shapeId, keep: [] });
		else if (edit.type === 'duplicate-shapes')
			for (const copy of edit.copies)
				targets.push({
					pageId: edit.pageId,
					shapeId: copy.shapeId,
					keep: [copy.shapeId, copy.newShapeId],
				});
	}
	return targets;
}

/**
 * Resizing, deleting or copying a shape that has a data graphic acts on the shape, not on the
 * callouts that follow it. The graphic is taken off the shape for the edit and written again
 * afterwards, laid out for the shape's new size, on the shape and on its copies.
 */
export async function stripDataGraphics(
	input: Uint8Array,
	edits: readonly VisioEdit[],
	limits: VisioPackageLimits,
	maxOutput: number,
	deadline: number,
): Promise<{ bytes: Uint8Array; restore: VisioSetDataGraphicEdit[] } | undefined> {
	const targets = targetsOf(edits);
	if (!targets.length) return undefined;
	const check = () => undefined;
	const { pkg, pages } = await openEditablePackage(input, limits, check);
	const strip: VisioSetDataGraphicEdit[] = [];
	const restore: VisioSetDataGraphicEdit[] = [];
	for (const target of targets) {
		const path = pages.get(target.pageId);
		if (
			!path ||
			strip.some((edit) => edit.pageId === target.pageId && edit.shapeId === target.shapeId)
		)
			continue;
		const owner = shapesOf(await visioXml(pkg, path, 'PageContents')).find(
			(shape) => attribute(shape, 'ID') === target.shapeId,
		);
		const items = owner && readDataGraphicItems(owner);
		if (!items) continue;
		strip.push({
			type: 'set-data-graphic',
			pageId: target.pageId,
			shapeId: target.shapeId,
			items: [],
		});
		for (const shapeId of target.keep)
			restore.push({ type: 'set-data-graphic', pageId: target.pageId, shapeId, items });
	}
	if (!strip.length) return undefined;
	const stripped = await editVsdxDataGraphics(input, strip, limits, maxOutput, deadline, check);
	return { bytes: stripped.bytes, restore };
}
