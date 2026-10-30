import { fail } from './package-common';
import type { VisioPartBox } from './edit-diagram-parts-commands';

/** Width of a lane's heading band and height of the frame's title bar, in inches. */
export const VISIO_SWIMLANE_HEADER = 0.5;
export const VISIO_SWIMLANE_TITLE = 0.35;
/** Thickness of a phase separator bar, in inches. */
export const VISIO_SEPARATOR_THICKNESS = 0.04;
export const VISIO_SWIMLANE_MIN = 0.75;
export const VISIO_SWIMLANE_STYLES = [1, 2, 3, 4] as const;

export type VisioSwimlaneDirection = 'horizontal' | 'vertical';

/** A cross-functional flowchart: a frame container and its lanes, stacked inside `box`. */
export interface VisioCreateSwimlanesEdit {
	type: 'create-swimlanes';
	pageId: string;
	frameId: string;
	laneIds: readonly string[];
	headings: readonly string[];
	/** The whole frame including the title bar, in drawing inches. */
	box: VisioPartBox;
	direction: VisioSwimlaneDirection;
	title: string;
	style?: number;
}
/** Insert a lane before or after `relativeTo`; later lanes shift and the frame grows. */
export interface VisioInsertSwimlaneEdit {
	type: 'insert-swimlane';
	pageId: string;
	frameId: string;
	shapeId: string;
	relativeTo: string;
	position: 'before' | 'after';
	heading: string;
	/** 'leave': the caller moves the shapes in shifted lanes (stencil shapes need it). */
	contents?: 'move' | 'leave';
}
/** Delete a lane together with the shapes in it; later lanes close up and the frame shrinks. */
export interface VisioDeleteSwimlaneEdit {
	type: 'delete-swimlane';
	pageId: string;
	frameId: string;
	laneId: string;
	/** 'leave': the caller moves the shapes in shifted lanes (stencil shapes need it). */
	contents?: 'move' | 'leave';
}
/** Change a lane's size along the stacking direction; lanes after it move and the frame follows. */
export interface VisioResizeSwimlaneEdit {
	type: 'resize-swimlane';
	pageId: string;
	frameId: string;
	laneId: string;
	size: number;
	/** 'leave': the caller moves the shapes in shifted lanes (stencil shapes need it). */
	contents?: 'move' | 'leave';
}
/** Add a phase separator bar across all lanes at `at` (drawing inches along the lane axis). */
export interface VisioInsertSeparatorEdit {
	type: 'insert-separator';
	pageId: string;
	frameId: string;
	shapeId: string;
	at?: number;
}
/** The Cross-Functional Flowchart tab's Design and Arrange options. */
export interface VisioSwimlaneOptionsEdit {
	type: 'set-swimlane-options';
	pageId: string;
	frameId: string;
	showTitle?: boolean;
	showSeparators?: boolean;
	rotateLabels?: boolean;
	style?: number;
	/** 'leave': the caller moves the shapes in shifted lanes (stencil shapes need it). */
	contents?: 'move' | 'leave';
}
export type VisioSwimlaneEdit =
	| VisioCreateSwimlanesEdit
	| VisioInsertSwimlaneEdit
	| VisioDeleteSwimlaneEdit
	| VisioResizeSwimlaneEdit
	| VisioInsertSeparatorEdit
	| VisioSwimlaneOptionsEdit;

const TYPES = new Set([
	'create-swimlanes',
	'insert-swimlane',
	'delete-swimlane',
	'resize-swimlane',
	'insert-separator',
	'set-swimlane-options',
]);
export const isVisioSwimlaneEdit = (edit: { type: string }): edit is VisioSwimlaneEdit =>
	TYPES.has(edit.type);

const id = (value: unknown): string => {
	if (typeof value !== 'string' || !/^[1-9]\d{0,9}$/.test(value) || Number(value) > 0xffffffff)
		fail('INVALID_EDIT', 'Swimlane shape IDs require canonical positive unsigned integers.');
	return value;
};
const text = (value: unknown): string => {
	if (
		typeof value !== 'string' ||
		value.length > 10_000 ||
		/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(value)
	)
		fail('INVALID_EDIT', 'Swimlane text must be short plain text.');
	return value;
};
const finite = (value: unknown, positive = false): number => {
	if (
		typeof value !== 'number' ||
		!Number.isFinite(value) ||
		Math.abs(value) > 1e6 ||
		(positive && value <= 0)
	)
		fail('INVALID_EDIT', 'Swimlane sizes must be finite inches.');
	return value;
};
const style = (value: unknown): number => {
	if (!(VISIO_SWIMLANE_STYLES as readonly unknown[]).includes(value))
		fail('INVALID_EDIT', 'Unknown cross-functional flowchart style.');
	return value as number;
};
const bool = (value: unknown): boolean => {
	if (typeof value !== 'boolean') fail('INVALID_EDIT', 'A swimlane option must be true or false.');
	return value;
};

const leave = (edit: { contents?: unknown }): { contents?: 'leave' } =>
	edit.contents === 'leave' ? { contents: 'leave' } : {};

/** Copy and validate a swimlane command; source admission stays authoritative. */
export function snapshotSwimlaneEdit(edit: VisioSwimlaneEdit): VisioSwimlaneEdit {
	if (typeof edit.pageId !== 'string' || !edit.pageId || edit.pageId.length > 256)
		fail('INVALID_EDIT', 'Invalid edit target.');
	const base = { pageId: edit.pageId, frameId: id(edit.frameId) };
	switch (edit.type) {
		case 'create-swimlanes': {
			if (!Array.isArray(edit.laneIds) || edit.laneIds.length < 1 || edit.laneIds.length > 100)
				fail('INVALID_EDIT', 'A cross-functional flowchart has 1 to 100 lanes.');
			const laneIds = Array.from(edit.laneIds, id);
			const headings = Array.from(edit.headings ?? [], text);
			if (new Set([...laneIds, base.frameId]).size !== laneIds.length + 1)
				fail('INVALID_EDIT', 'The frame and its lanes need distinct IDs.');
			if (edit.direction !== 'horizontal' && edit.direction !== 'vertical')
				fail('INVALID_EDIT', 'Unknown swimlane direction.');
			const b = edit.box ?? ({} as VisioPartBox);
			return {
				type: edit.type,
				...base,
				laneIds,
				headings,
				box: {
					x: finite(b.x),
					y: finite(b.y),
					width: finite(b.width, true),
					height: finite(b.height, true),
				},
				direction: edit.direction,
				title: text(edit.title),
				...(edit.style === undefined ? {} : { style: style(edit.style) }),
			};
		}
		case 'insert-swimlane':
			if (edit.position !== 'before' && edit.position !== 'after')
				fail('INVALID_EDIT', 'A lane goes before or after another lane.');
			return {
				type: edit.type,
				...base,
				shapeId: id(edit.shapeId),
				relativeTo: id(edit.relativeTo),
				position: edit.position,
				heading: text(edit.heading),
				...leave(edit),
			};
		case 'delete-swimlane':
			return { type: edit.type, ...base, laneId: id(edit.laneId), ...leave(edit) };
		case 'resize-swimlane':
			return {
				type: edit.type,
				...base,
				laneId: id(edit.laneId),
				size: finite(edit.size, true),
				...leave(edit),
			};
		case 'insert-separator':
			return {
				type: edit.type,
				...base,
				shapeId: id(edit.shapeId),
				...(edit.at === undefined ? {} : { at: finite(edit.at) }),
			};
		default:
			return {
				type: 'set-swimlane-options',
				...base,
				...(edit.showTitle === undefined ? {} : { showTitle: bool(edit.showTitle) }),
				...(edit.showSeparators === undefined ? {} : { showSeparators: bool(edit.showSeparators) }),
				...(edit.rotateLabels === undefined ? {} : { rotateLabels: bool(edit.rotateLabels) }),
				...(edit.style === undefined ? {} : { style: style(edit.style) }),
				...leave(edit),
			};
	}
}
