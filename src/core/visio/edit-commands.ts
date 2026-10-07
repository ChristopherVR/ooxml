import { fail } from './package-common';

export interface VisioTextEdit {
	type: 'replace-plain-text';
	pageId: string;
	shapeId: string;
	text: string;
}
interface Target {
	pageId: string;
	shapeId: string;
}
/** Drawing inches; bottom-left origin, up-positive rotation pin. */
export type VisioGeometryEdit =
	| (Target & {
			type: 'create-rectangle';
			x: number;
			y: number;
			width: number;
			height: number;
			text?: string;
	  })
	| (Target & { type: 'move-shape'; x: number; y: number })
	| (Target & { type: 'resize-shape'; width: number; height: number })
	| (Target & { type: 'move-line-endpoint'; endpoint: 'begin' | 'end'; x: number; y: number })
	| (Target & { type: 'delete-shape' });
/** Insert a blank foreground page after an existing page, copying its PageSheet settings. */
export interface VisioPageInsert {
	type: 'insert-page';
	pageId: string;
	afterPageId: string;
	name: string;
}
/** Move an existing page to a zero-based index in the complete page list. */
export interface VisioPageReorder {
	type: 'reorder-page';
	pageId: string;
	index: number;
}
/** Rename the local page name; an already custom universal name is preserved. */
export interface VisioPageRename {
	type: 'rename-page';
	pageId: string;
	name: string;
}
/** Delete a page, preserving remaining page names (native Page.Delete(0)). */
export interface VisioPageDelete {
	type: 'delete-page';
	pageId: string;
}
export type VisioPageEdit = VisioPageInsert | VisioPageReorder | VisioPageRename | VisioPageDelete;
export type VisioEdit = VisioTextEdit | VisioGeometryEdit | VisioPageEdit;
export const isVisioPageEdit = (edit: VisioEdit): edit is VisioPageEdit =>
	edit.type === 'insert-page' ||
	edit.type === 'reorder-page' ||
	edit.type === 'rename-page' ||
	edit.type === 'delete-page';

/** Potential direct changes used by both package and master dependency admission. */
export function geometryChangedCells(edit: VisioGeometryEdit): string[] {
	if (edit.type === 'delete-shape') return [];
	if (edit.type === 'move-line-endpoint') {
		const prefix = edit.endpoint === 'begin' ? 'Begin' : 'End';
		return [`${prefix}X`, `${prefix}Y`];
	}
	if (edit.type === 'move-shape') return ['PinX', 'PinY'];
	if (edit.type === 'resize-shape') return ['Width', 'Height'];
	return ['PinX', 'PinY', 'Width', 'Height'];
}

export function snapshotVisioEdits(
	edits: readonly VisioEdit[],
	maxEdits: number,
	maxText: number,
): VisioEdit[] {
	if (!Array.isArray(edits) || edits.length > maxEdits)
		fail('LIMIT_EDITS', 'Edit command count exceeds limit.');
	let textLength = 0;
	const text = (value: unknown): string => {
		if (typeof value !== 'string') fail('INVALID_EDIT', 'Edit text must be a string.');
		textLength += value.length;
		if (textLength > maxText) fail('LIMIT_EDIT_TEXT', 'Replacement text exceeds aggregate limit.');
		if (
			/[\u0000-\u0008\u000b\u000c\u000d\u000e-\u001f\ufffe\uffff]|[\ud800-\udbff](?![\udc00-\udfff])|(?<![\ud800-\udbff])[\udc00-\udfff]/u.test(
				value,
			)
		)
			fail('INVALID_EDIT_TEXT', 'Replacement text contains invalid XML characters.');
		return value;
	};
	const numeric = (value: unknown, dimension = false, allowZero = false): number => {
		if (
			typeof value !== 'number' ||
			!Number.isFinite(value) ||
			Math.abs(value) > 1e6 ||
			(dimension && (value < 0 || (!allowZero && value === 0)))
		)
			fail(
				'INVALID_EDIT',
				'Geometry requires finite drawing inches within limits and positive dimensions.',
			);
		return value;
	};
	return Array.from(edits, (edit) => {
		if (!edit || typeof edit.pageId !== 'string' || !edit.pageId || edit.pageId.length > 256)
			fail('INVALID_EDIT', 'Invalid edit target.');
		if (edit.type === 'delete-page') return { type: edit.type, pageId: edit.pageId };
		if (edit.type === 'reorder-page') {
			if (!Number.isSafeInteger(edit.index) || edit.index < 0 || edit.index > 1_000_000)
				fail('INVALID_EDIT', 'Page order requires a bounded zero-based integer index.');
			return { type: edit.type, pageId: edit.pageId, index: edit.index };
		}
		if (edit.type === 'insert-page') {
			const name = text(edit.name);
			if (
				!/^(0|[1-9]\d{0,9})$/.test(edit.pageId) ||
				Number(edit.pageId) > 4294967295 ||
				typeof edit.afterPageId !== 'string' ||
				!edit.afterPageId ||
				edit.afterPageId.length > 256 ||
				!name.trim() ||
				name.length > 256 ||
				/[\t\n]/.test(name)
			)
				fail('INVALID_EDIT', 'Invalid new page ID, name or insertion target.');
			return { type: edit.type, pageId: edit.pageId, afterPageId: edit.afterPageId, name };
		}
		if (edit.type === 'rename-page') {
			const name = text(edit.name);
			if (!name.trim() || name.length > 255 || /[\t\n]/.test(name))
				fail('INVALID_EDIT', 'Invalid page name.');
			return { type: edit.type, pageId: edit.pageId, name };
		}
		if (typeof edit.shapeId !== 'string' || !edit.shapeId || edit.shapeId.length > 256)
			fail('INVALID_EDIT', 'Invalid edit shape target.');
		const target = { pageId: edit.pageId, shapeId: edit.shapeId };
		if (edit.type === 'replace-plain-text')
			return { ...target, type: edit.type, text: text(edit.text) };
		if (!/^[1-9]\d{0,9}$/.test(edit.shapeId) || Number(edit.shapeId) > 4294967295)
			fail('INVALID_EDIT', 'Geometry shape IDs must be canonical positive unsigned integers.');
		switch (edit.type) {
			case 'create-rectangle':
				return {
					...target,
					type: edit.type,
					x: numeric(edit.x),
					y: numeric(edit.y),
					width: numeric(edit.width, true),
					height: numeric(edit.height, true),
					...(edit.text === undefined ? {} : { text: text(edit.text) }),
				};
			case 'move-shape':
				return { ...target, type: edit.type, x: numeric(edit.x), y: numeric(edit.y) };
			case 'move-line-endpoint':
				if (edit.endpoint !== 'begin' && edit.endpoint !== 'end')
					fail('INVALID_EDIT', 'A line endpoint must be begin or end.');
				return {
					...target,
					type: edit.type,
					endpoint: edit.endpoint,
					x: numeric(edit.x),
					y: numeric(edit.y),
				};
			case 'resize-shape':
				return {
					...target,
					type: edit.type,
					width: numeric(edit.width, true),
					height: numeric(edit.height, true, true),
				};
			case 'delete-shape':
				return { ...target, type: edit.type };
			default:
				return fail('INVALID_EDIT', 'Unsupported edit command.');
		}
	});
}
