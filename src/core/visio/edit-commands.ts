import { fail } from './package-common.js';

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
	| (Target & { type: 'delete-shape' });
export type VisioEdit = VisioTextEdit | VisioGeometryEdit;

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
	const numeric = (value: unknown, dimension = false): number => {
		if (
			typeof value !== 'number' ||
			!Number.isFinite(value) ||
			Math.abs(value) > 1e6 ||
			(dimension && value <= 0)
		)
			fail(
				'INVALID_EDIT',
				'Geometry requires finite drawing inches within limits and positive dimensions.',
			);
		return value;
	};
	return Array.from(edits, (edit) => {
		if (
			!edit ||
			typeof edit.pageId !== 'string' ||
			!edit.pageId ||
			edit.pageId.length > 256 ||
			typeof edit.shapeId !== 'string' ||
			!edit.shapeId ||
			edit.shapeId.length > 256
		)
			fail('INVALID_EDIT', 'Invalid edit target.');
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
			case 'resize-shape':
				return {
					...target,
					type: edit.type,
					width: numeric(edit.width, true),
					height: numeric(edit.height, true),
				};
			case 'delete-shape':
				return { ...target, type: edit.type };
			default:
				return fail('INVALID_EDIT', 'Unsupported edit command.');
		}
	});
}
