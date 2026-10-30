import { children } from './sheet';
import { fail } from './package-common';
import { setCell } from './edit-geometry-cells';
import { deleteVisioShapes } from './edit-delete';
import { insertVisioContainer } from './edit-diagram-parts';
import { boundsOf, type Box } from './edit-container-fit';
import {
	containerClosure,
	shapeIdOf,
	topLevelShapes,
	userValue,
	writeContainerMembers,
	writeListMembers,
} from './edit-container-members';
import { turnFlowchart } from './edit-swimlane-turn';
import {
	VISIO_CONTAINER_MARGIN_ROW,
	VISIO_CONTAINER_RESIZE_ROW,
} from './edit-diagram-parts-commands';
import { VISIO_SWIMLANE_FRAME_ROW } from './shape-structure';
import {
	VISIO_SEPARATOR_THICKNESS,
	VISIO_SWIMLANE_MIN,
	VISIO_SWIMLANE_TITLE,
	type VisioCreateSwimlanesEdit,
	type VisioDeleteSwimlaneEdit,
	type VisioInsertSeparatorEdit,
	type VisioInsertSwimlaneEdit,
	type VisioResizeSwimlaneEdit,
	type VisioSwimlaneEdit,
	type VisioSwimlaneOptionsEdit,
} from './edit-swimlane-commands';
import {
	VISIO_SWIMLANE_TITLE_ROW,
	applyFrameTitle,
	applyLaneLabel,
	applyLaneStyle,
	applySeparatorLabel,
	createLaneShape,
	createSeparatorShape,
	setTextCell,
	setUserRow,
} from './edit-swimlane-shapes';
import {
	ensureLists,
	frameOf,
	laneOf,
	lanesOf,
	optionsOf,
	orderSeparators,
	place,
	relayout,
	sendBehind,
	separatorsOf,
	type Scope,
} from './edit-swimlane-layout';

const flagText = (value: boolean) => (value ? '1' : '0');

function createSwimlanes(scope: Scope, edit: VisioCreateSwimlanesEdit): void {
	const { root, document, check } = scope;
	const horizontal = edit.direction === 'horizontal';
	const style = edit.style ?? 1;
	const { box } = edit;
	const count = edit.laneIds.length;
	insertVisioContainer(
		root,
		document,
		{
			type: 'insert-container',
			pageId: edit.pageId,
			shapeId: edit.frameId,
			memberIds: [],
			style: 'plain',
			heading: edit.title,
			box,
		},
		check,
	);
	const frame = topLevelShapes(root).find((shape) => shapeIdOf(shape) === edit.frameId)!;
	setCell(frame, 'FillPattern', 0);
	setTextCell(frame, 'LineColor', '#7F7F7F');
	applyFrameTitle(frame, true);
	for (const [name, value, unit] of [
		[VISIO_CONTAINER_MARGIN_ROW, '0', 'IN'],
		[VISIO_CONTAINER_RESIZE_ROW, '0', 'STR'],
		[VISIO_SWIMLANE_FRAME_ROW, '1', 'STR'],
		// Visio's own categories and flowchart rows (numbers carry no unit, as Visio writes them).
		['msvShapeCategories', 'CFF Container;DoNotContain', 'STR'],
		['msvSDContainerRequiredCategories', 'CFF List', 'STR'],
		['msvSDContainerLocked', '1', 'BOOL'],
		['numLanes', String(count), ''],
		['visCFFStyle', String(style), ''],
		['visShowTitle', '1', ''],
		['visShowPhase', '1', ''],
		['visMargins', '0', 'IN'],
		['CFFVertical', flagText(!horizontal), ''],
		['visRotateLabel', '0', ''],
		['RTL', '0', ''],
		['visCFFHeadingPos', '2', ''],
		['visVersion', '15', ''],
	] as const)
		setUserRow(frame, name, value, unit);
	const title = VISIO_SWIMLANE_TITLE;
	const left = box.x - box.width / 2;
	const top = box.y + box.height / 2 - title;
	const lanes = edit.laneIds.map((id, index) => {
		const heading = edit.headings[index] ?? `Function ${index + 1}`;
		const laneBox = horizontal
			? {
					x: box.x,
					y: top - ((index + 0.5) * (box.height - title)) / count,
					width: box.width,
					height: (box.height - title) / count,
				}
			: {
					x: left + ((index + 0.5) * box.width) / count,
					y: box.y - title / 2,
					width: box.width / count,
					height: box.height - title,
				};
		return createLaneShape(root, document, id, laneBox, heading, edit.direction, false, style);
	});
	writeContainerMembers(root, frame, edit.laneIds);
	const { swim, phase } = ensureLists(scope, frame, {
		...(edit.listId ? { swim: edit.listId } : {}),
		...(edit.phaseListId ? { phase: edit.phaseListId } : {}),
	});
	sendBehind(root, [swim, phase, frame, ...lanes]);
}

function insertSwimlane(scope: Scope, edit: VisioInsertSwimlaneEdit): void {
	const frame = frameOf(scope, edit.frameId);
	ensureLists(scope, frame);
	const options = optionsOf(frame);
	const lanes = lanesOf(scope, frame, options.direction);
	const reference = laneOf(scope, frame, edit.relativeTo, options.direction);
	const old = boundsOf([reference]);
	if (!old) fail('UNSUPPORTED_GEOMETRY_EDIT', 'The lane has no usable bounds.');
	const lane = createLaneShape(
		scope.root,
		scope.document,
		edit.shapeId,
		{
			x: (old.minX + old.maxX) / 2,
			y: (old.minY + old.maxY) / 2,
			width: old.maxX - old.minX,
			height: old.maxY - old.minY,
		},
		edit.heading,
		options.direction,
		options.rotate,
		options.style,
	);
	frame.parentNode!.insertBefore(lane, frame.nextSibling);
	const order = [...lanes];
	order.splice(order.indexOf(reference) + (edit.position === 'after' ? 1 : 0), 0, lane);
	relayout(scope, frame, order);
}

function deleteSwimlane(
	scope: Scope,
	edit: VisioDeleteSwimlaneEdit,
	instances: ReadonlySet<Element>,
): void {
	const frame = frameOf(scope, edit.frameId);
	ensureLists(scope, frame);
	const options = optionsOf(frame);
	if (lanesOf(scope, frame, options.direction).length < 2)
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'The last lane of a flowchart cannot be deleted.');
	const lane = laneOf(scope, frame, edit.laneId, options.direction);
	const present = new Set(topLevelShapes(scope.root).map(shapeIdOf));
	// The lane's shapes go with it in the same step, stencil instances included (the transaction
	// admitted them: `instances`); `contents` only decides who moves the lanes that close up.
	const ids = [edit.laneId, ...containerClosure(scope.root, lane)].filter((id) => present.has(id));
	deleteVisioShapes(
		scope.roots,
		scope.document,
		ids.map((shapeId) => ({ type: 'delete-shape' as const, pageId: edit.pageId, shapeId })),
		scope.check,
		instances,
	);
	relayout(scope, frame);
}

/** The shapes a delete-swimlane command removes with its lane (for the transaction's admission). */
export function swimlaneDeleteTargets(
	roots: ReadonlyMap<string, Element>,
	command: { type: string },
): { pageId: string; shapeId: string }[] {
	if (command.type !== 'delete-swimlane') return [];
	const edit = command as VisioDeleteSwimlaneEdit;
	const root = roots.get(edit.pageId);
	const lane = root && topLevelShapes(root).find((shape) => shapeIdOf(shape) === edit.laneId);
	if (!root || !lane) return [];
	return [...containerClosure(root, lane)].map((shapeId) => ({ pageId: edit.pageId, shapeId }));
}

function resizeSwimlane(scope: Scope, edit: VisioResizeSwimlaneEdit): void {
	if (edit.size < VISIO_SWIMLANE_MIN)
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'A lane cannot be narrower than three quarters of an inch.');
	const frame = frameOf(scope, edit.frameId);
	ensureLists(scope, frame);
	const options = optionsOf(frame);
	const lane = laneOf(scope, frame, edit.laneId, options.direction);
	const old = boundsOf([lane])!;
	const target: Box =
		options.direction === 'horizontal'
			? { ...old, minY: old.maxY - edit.size }
			: { ...old, maxX: old.minX + edit.size };
	place(scope, lane, target);
	relayout(scope, frame);
}

function insertSeparator(scope: Scope, edit: VisioInsertSeparatorEdit): void {
	const frame = frameOf(scope, edit.frameId);
	const { phase } = ensureLists(scope, frame);
	const options = optionsOf(frame);
	const lanes = lanesOf(scope, frame, options.direction);
	const area = boundsOf(lanes);
	if (!area) fail('UNSUPPORTED_GEOMETRY_EDIT', 'The flowchart has no lanes.');
	const horizontal = options.direction === 'horizontal';
	const thick = VISIO_SEPARATOR_THICKNESS;
	const low = horizontal ? area.minX : area.minY;
	const high = horizontal ? area.maxX : area.maxY;
	const at = Math.min(Math.max(edit.at ?? (low + high) / 2, low), high);
	const bar = createSeparatorShape(
		scope.root,
		scope.document,
		edit.shapeId,
		horizontal
			? {
					x: at,
					y: (area.minY + area.maxY) / 2,
					width: thick,
					height: area.maxY - area.minY,
				}
			: {
					x: (area.minX + area.maxX) / 2,
					y: at,
					width: area.maxX - area.minX,
					height: thick,
				},
		options.style,
		options.separators,
		options.direction,
		edit.label ?? `Phase ${separatorsOf(scope, frame).length + 1}`,
	);
	const last = lanes[lanes.length - 1]!;
	last.parentNode!.insertBefore(bar, last.nextSibling);
	writeListMembers(
		scope.root,
		phase,
		orderSeparators([...separatorsOf(scope, frame), bar], options.direction).map(shapeIdOf),
	);
}

function setOptions(scope: Scope, edit: VisioSwimlaneOptionsEdit): void {
	const frame = frameOf(scope, edit.frameId);
	const before = optionsOf(frame);
	if (edit.showTitle !== undefined) setUserRow(frame, 'visShowTitle', flagText(edit.showTitle));
	if (edit.showSeparators !== undefined)
		setUserRow(frame, 'visShowPhase', flagText(edit.showSeparators));
	if (edit.rotateLabels !== undefined)
		setUserRow(frame, 'visRotateLabel', flagText(edit.rotateLabels));
	if (edit.style !== undefined) setUserRow(frame, 'visCFFStyle', String(edit.style));
	if (edit.direction !== undefined && edit.direction !== before.direction)
		turnFlowchart(scope, frame, edit.direction);
	const options = optionsOf(frame);
	const lanes = lanesOf(scope, frame, options.direction);
	const separators = separatorsOf(scope, frame);
	if (options.title !== before.title) {
		const text = children(frame, 'Text')[0];
		if (text) {
			if (!options.title) {
				setUserRow(frame, VISIO_SWIMLANE_TITLE_ROW, text.textContent ?? '');
				text.textContent = '';
			} else text.textContent = userValue(frame, VISIO_SWIMLANE_TITLE_ROW) ?? '';
		}
		applyFrameTitle(frame, options.title);
	}
	if (options.rotate !== before.rotate)
		for (const lane of lanes) applyLaneLabel(lane, options.direction, options.rotate);
	if (options.style !== before.style) {
		for (const lane of lanes) applyLaneStyle(lane, options.style, false);
		applyLaneStyle(frame, options.style, true);
		for (const bar of separators) setTextCell(bar, 'FillForegnd', styleLine(options.style));
	}
	for (const bar of separators) {
		setCell(bar, 'FillPattern', options.separators ? 1 : 0);
		applySeparatorLabel(bar, options.direction, options.separators);
	}
	if (options.title !== before.title) relayout(scope, frame);
}
const styleLine = (style: number) =>
	['#C99A55', '#5B9BD5', '#7F7F7F', '#6AA84F'][style - 1] ?? '#C99A55';

/** Apply a cross-functional flowchart command to its page; throws a VisioPackageError on refusal. */
export function applySwimlaneEdit(
	roots: ReadonlyMap<string, Element>,
	document: Element,
	edit: VisioSwimlaneEdit,
	check: () => void,
	instances: ReadonlySet<Element> = new Set(),
): void {
	check();
	const root = roots.get(edit.pageId);
	if (!root) fail('EDIT_TARGET_NOT_FOUND', 'Page does not exist.');
	const contents = 'contents' in edit && edit.contents === 'leave' ? 'leave' : 'move';
	const scope: Scope = { roots, document, root, pageId: edit.pageId, check, contents };
	if (edit.type === 'create-swimlanes') createSwimlanes(scope, edit);
	else if (edit.type === 'insert-swimlane') insertSwimlane(scope, edit);
	else if (edit.type === 'delete-swimlane') deleteSwimlane(scope, edit, instances);
	else if (edit.type === 'resize-swimlane') resizeSwimlane(scope, edit);
	else if (edit.type === 'insert-separator') insertSeparator(scope, edit);
	else setOptions(scope, edit);
}
