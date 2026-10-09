import { fail } from './package-common';

/** Group top-level local shapes under a new Type="Group" sheet (native Group, Ctrl+Shift+G). */
export interface VisioGroupShapesEdit {
	type: 'group-shapes';
	pageId: string;
	/** The new group's sheet ID; it must be unused on the page. */
	shapeId: string;
	/** Two or more top-level members; stacking order inside the group follows the page order. */
	memberIds: readonly string[];
}
/** Dissolve a top-level group, returning its direct members to page coordinates. */
export interface VisioUngroupShapeEdit {
	type: 'ungroup-shape';
	pageId: string;
	shapeId: string;
}
export type VisioGroupEdit = VisioGroupShapesEdit | VisioUngroupShapeEdit;

export const isVisioGroupEdit = (edit: { type: string }): edit is VisioGroupEdit =>
	edit.type === 'group-shapes' || edit.type === 'ungroup-shape';

const sheetId = (value: unknown): string => {
	if (typeof value !== 'string' || !/^[1-9]\d{0,9}$/.test(value) || Number(value) > 0xffffffff)
		fail('INVALID_EDIT', 'Group shape IDs require canonical positive unsigned integers.');
	return value;
};

/** Copy and validate a grouping command; source admission remains authoritative. */
export function snapshotGroupEdit(edit: VisioGroupEdit): VisioGroupEdit {
	if (typeof edit.pageId !== 'string' || !edit.pageId || edit.pageId.length > 256)
		fail('INVALID_EDIT', 'Invalid edit target.');
	const shapeId = sheetId(edit.shapeId);
	if (edit.type === 'ungroup-shape') return { type: edit.type, pageId: edit.pageId, shapeId };
	if (!Array.isArray(edit.memberIds) || edit.memberIds.length < 2 || edit.memberIds.length > 1000)
		fail('INVALID_EDIT', 'Grouping requires between two and 1000 member shapes.');
	const memberIds = Array.from(edit.memberIds, sheetId);
	if (new Set(memberIds).size !== memberIds.length || memberIds.includes(shapeId))
		fail('INVALID_EDIT', 'Group members must be unique and differ from the new group ID.');
	return { type: edit.type, pageId: edit.pageId, shapeId, memberIds };
}
