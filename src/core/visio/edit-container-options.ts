import { attribute, children } from './sheet';
import { fail } from './package-common';
import {
	fitVisioContainer,
	CONTAINER_RESIZE_NONE,
	CONTAINER_RESIZE_EXPAND,
	CONTAINER_RESIZE_TIGHT,
} from './edit-container-fit';
import { isContainerSheet, shapeIdOf, topLevelShapes } from './edit-container-members';
import {
	VISIO_CONTAINER_MARGIN_ROW,
	VISIO_CONTAINER_RESIZE_ROW,
} from './edit-diagram-parts-commands';

/** Visio's Lock Container flag (a BOOL User row); membership cannot change while it is set. */
export const VISIO_CONTAINER_LOCKED_ROW = 'msvSDContainerLocked';

/** Container Format > Fit to Contents: the frame takes the members' box, whatever its resize mode. */
export interface VisioContainerFitEdit {
	type: 'fit-container';
	pageId: string;
	shapeId: string;
}
/**
 * Container Format options: the margin in inches (msvSDContainerMargin), the resize mode
 * (msvSDContainerResize: 0 no automatic resize, 1 expand as needed, 2 always fit to contents) and
 * the lock (msvSDContainerLocked). A changed margin or mode refits the container at once.
 */
export interface VisioContainerOptionsEdit {
	type: 'set-container-options';
	pageId: string;
	shapeId: string;
	margin?: number;
	resize?: 0 | 1 | 2;
	locked?: boolean;
}
export type VisioContainerEdit = VisioContainerFitEdit | VisioContainerOptionsEdit;

export const isVisioContainerEdit = (edit: { type: string }): edit is VisioContainerEdit =>
	edit.type === 'fit-container' || edit.type === 'set-container-options';

const MAX_MARGIN = 10;

/** Copy and validate a container option edit; the target container is checked when applied. */
export function snapshotContainerEdit(edit: VisioContainerEdit): VisioContainerEdit {
	if (typeof edit.pageId !== 'string' || !edit.pageId || edit.pageId.length > 256)
		fail('INVALID_EDIT', 'Invalid edit target.');
	const shapeId = edit.shapeId;
	if (typeof shapeId !== 'string' || !/^[1-9]\d{0,9}$/.test(shapeId))
		fail('INVALID_EDIT', 'Container shape IDs require canonical positive unsigned integers.');
	if (edit.type === 'fit-container') return { type: edit.type, pageId: edit.pageId, shapeId };
	const result: VisioContainerOptionsEdit = { type: edit.type, pageId: edit.pageId, shapeId };
	if (edit.margin !== undefined) {
		const margin = edit.margin;
		if (typeof margin !== 'number' || !Number.isFinite(margin) || margin < 0 || margin > MAX_MARGIN)
			fail('INVALID_EDIT', 'A container margin is a finite length from 0 to 10 inches.');
		result.margin = margin;
	}
	if (edit.resize !== undefined) {
		if (
			![CONTAINER_RESIZE_NONE, CONTAINER_RESIZE_EXPAND, CONTAINER_RESIZE_TIGHT].includes(
				edit.resize,
			)
		)
			fail('INVALID_EDIT', 'Unknown container resize mode.');
		result.resize = edit.resize;
	}
	if (edit.locked !== undefined) {
		if (typeof edit.locked !== 'boolean')
			fail('INVALID_EDIT', 'A container lock is true or false.');
		result.locked = edit.locked;
	}
	if (result.margin === undefined && result.resize === undefined && result.locked === undefined)
		fail('INVALID_EDIT', 'A container option edit changes at least one option.');
	return result;
}

/** Set (or add) a User row's Value cell; the stored value is explicit, with no formula. */
function setUserValue(shape: Element, name: string, value: string, unit?: string): void {
	const doc = shape.ownerDocument!;
	const make = (tag: string) => doc.createElementNS(shape.namespaceURI, tag);
	let section = children(shape, 'Section').find((node) => attribute(node, 'N') === 'User');
	if (!section) {
		section = make('Section');
		section.setAttribute('N', 'User');
		shape.appendChild(section);
	}
	let row = children(section, 'Row').find((node) => attribute(node, 'N') === name);
	if (!row) {
		row = make('Row');
		row.setAttribute('N', name);
		section.appendChild(row);
	}
	let cell = children(row, 'Cell').find((node) => attribute(node, 'N') === 'Value');
	if (!cell) {
		cell = make('Cell');
		cell.setAttribute('N', 'Value');
		row.appendChild(cell);
	}
	cell.setAttribute('V', value);
	cell.removeAttribute('F');
	if (unit) cell.setAttribute('U', unit);
	else cell.removeAttribute('U');
}

/**
 * Apply a container edit to the page roots. Returns whether the page changed. A container that
 * the geometry editor does not admit keeps its size; its options are still written.
 */
export function applyContainerEdit(
	roots: ReadonlyMap<string, Element>,
	document: Element,
	edit: VisioContainerEdit,
	check: () => void,
): boolean {
	check();
	const root = roots.get(edit.pageId);
	if (!root) fail('EDIT_TARGET_NOT_FOUND', 'Page does not exist.');
	const container = topLevelShapes(root).find((shape) => shapeIdOf(shape) === edit.shapeId);
	if (!container || !isContainerSheet(container))
		fail('EDIT_TARGET_NOT_FOUND', 'A top-level container shape is required.');
	if (edit.type === 'fit-container')
		return fitVisioContainer(roots, document, edit.pageId, container, check, true);
	if (edit.margin !== undefined)
		setUserValue(container, VISIO_CONTAINER_MARGIN_ROW, String(edit.margin), 'IN');
	if (edit.resize !== undefined)
		setUserValue(container, VISIO_CONTAINER_RESIZE_ROW, String(edit.resize));
	if (edit.locked !== undefined)
		setUserValue(container, VISIO_CONTAINER_LOCKED_ROW, edit.locked ? '1' : '0', 'BOOL');
	if (edit.margin !== undefined || edit.resize !== undefined)
		fitVisioContainer(roots, document, edit.pageId, container, check);
	return true;
}
