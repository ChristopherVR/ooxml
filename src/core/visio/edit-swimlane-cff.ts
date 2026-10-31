import { attribute, children } from './sheet';
import { fail } from './package-common';
import { boundsOf, type Box } from './edit-container-fit';
import { shapeIdOf, userValue, writeListMembers } from './edit-container-members';
import { cloneNativeLane, headingShape } from './edit-swimlane-clone';
import { setUserRow } from './edit-swimlane-shapes';
import type { VisioInsertSeparatorEdit, VisioSwimlaneOptionsEdit } from './edit-swimlane-commands';
import { recalculateMasterInstances } from './cff-recalc';
import {
	PHASE_BAND,
	TITLE_BAND,
	ensureLists,
	headingBand,
	lanesOf,
	masterRole,
	optionsOf,
	orderSeparators,
	place,
	relayout,
	separatorsOf,
	type Scope,
} from './edit-swimlane-layout';

/** The narrowest phase Insert Phase leaves on either side of the new separator, in inches. */
export const CFF_MIN_PHASE = 0.4;
/** Whether a frame is an instance of Visio's own CFF Container master (a chart Visio built). */
export const isNativeFrame = (scope: Scope, frame: Element): boolean =>
	!scope.masters.empty && masterRole(scope, frame) === 'frame';

const flagText = (value: boolean) => (value ? '1' : '0');

/**
 * Insert Phase on a chart Visio built. Visio's Phase List lays its separators out edge to edge,
 * each as wide as its phase with the dividing line at its right edge, so a new separator splits
 * the phase it lands in: the copy takes the left part, the separator that was there keeps the
 * right part (observed in Visio 16: a Separator dropped at x = 8 in the template's one phase got
 * the width up to 8 and the old one the rest, and the Phase List named them new first). Only the
 * first phase carries the heading extension over the lane headings (`FirstHeadingExtend`).
 */
export function insertNativePhase(
	scope: Scope,
	frame: Element,
	edit: VisioInsertSeparatorEdit,
): void {
	if (optionsOf(frame).direction !== 'horizontal')
		fail(
			'UNSUPPORTED_GEOMETRY_EDIT',
			'Phases of a vertical chart Visio built are not editable yet: the list layout was only confirmed against Visio for horizontal charts.',
		);
	const { phase } = ensureLists(scope, frame);
	const bars = orderSeparators(separatorsOf(scope, frame), 'horizontal');
	const boxes = bars.map((bar) => boundsOf([bar]));
	if (!bars.length || boxes.some((box) => !box))
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'The flowchart has no usable phase to split.');
	const phases = boxes as Box[];
	const target = edit.at;
	let index = 0;
	if (target === undefined) {
		// Without a drop point, split the widest phase in the middle.
		phases.forEach((box, i) => {
			if (box.maxX - box.minX > phases[index]!.maxX - phases[index]!.minX) index = i;
		});
	} else {
		index = phases.findIndex((box) => target <= box.maxX);
		if (index < 0) index = phases.length - 1;
	}
	const split = phases[index]!;
	if (split.maxX - split.minX < 2 * CFF_MIN_PHASE)
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'No phase is wide enough to hold another separator.');
	const wanted = target ?? (split.minX + split.maxX) / 2;
	const at = Math.min(Math.max(wanted, split.minX + CFF_MIN_PHASE), split.maxX - CFF_MIN_PHASE);
	const old = bars[index]!;
	const bar = cloneNativeLane(scope.root, old, edit.shapeId, edit.label);
	const order = [...bars];
	order.splice(index, 0, bar);
	// The first phase alone extends over the lane headings; the one pushed to second place loses it.
	if (index === 0) dropFirstExtension(old);
	place(scope, bar, { ...split, maxX: at });
	place(scope, old, { ...split, minX: at });
	writeListMembers(scope.root, phase, order.map(shapeIdOf));
}

function dropFirstExtension(separator: Element): void {
	const heading = headingShape(separator);
	if (!heading) return;
	for (const section of children(heading, 'Section'))
		if (attribute(section, 'N') === 'User')
			for (const row of children(section, 'Row'))
				if (attribute(row, 'N') === 'FirstHeadingExtend') section.removeChild(row);
}

/** `visCFFSettings` is the string the add-in reads back: its numbers follow the frame's cells. */
function refreshSettings(scope: Scope, frame: Element, swim: Element): void {
	const values: Record<string, string> = {
		stg1: userValue(frame, 'visCFFStyle') ?? '1',
		stg2: userValue(frame, 'visShowPhase') ?? '1',
		stg3: userValue(frame, 'visShowTitle') ?? '1',
		stg6: userValue(swim, 'visRotateLabel') ?? '0',
	};
	for (const lane of Array.from(scope.root.getElementsByTagName('Shape'))) {
		for (const section of children(lane, 'Section'))
			if (attribute(section, 'N') === 'User')
				for (const row of children(section, 'Row'))
					if (attribute(row, 'N') === 'visCFFSettings')
						for (const cell of children(row, 'Cell'))
							if (attribute(cell, 'N') === 'Value')
								cell.setAttribute(
									'V',
									(attribute(cell, 'V') ?? '').replace(
										/\/(stg[1236])=[^,]*/g,
										(_, key: string) => `/${key}=${values[key]}`,
									),
								);
	}
}

/**
 * Chart options on a chart Visio built: style, rotated lane labels, the title bar and the phase
 * headings. They are the flowchart's own User cells (`visCFFStyle`, `visShowTitle`,
 * `visShowPhase` on the frame, `visRotateLabel` on the Swimlane List) that Visio's masters read
 * through formulas, so the cells are written as Visio writes them and every instance is
 * recalculated. Geometry follows what Visio 16 did to the template chart when the cells changed:
 * hiding the title bar lowers the frame's top by the bar and nothing else moves; hiding the
 * phase headings takes their band out from under the frame's top, so the lanes (with their
 * shapes) and the lists rise by the band and the frame and Phase List shrink from below.
 * Showing them reverses it.
 */
export function setNativeOptions(
	scope: Scope,
	frame: Element,
	edit: VisioSwimlaneOptionsEdit,
): void {
	const before = optionsOf(frame);
	if (edit.direction !== undefined && edit.direction !== before.direction)
		fail(
			'UNSUPPORTED_GEOMETRY_EDIT',
			'Direction of a chart Visio built cannot be changed yet: Visio swaps the flowchart masters for vertical ones, which this editor does not copy.',
		);
	const titleChanges = edit.showTitle !== undefined && edit.showTitle !== before.title;
	const phaseChanges =
		edit.showSeparators !== undefined && edit.showSeparators !== before.separators;
	if (before.direction !== 'horizontal' && (titleChanges || phaseChanges))
		fail(
			'UNSUPPORTED_GEOMETRY_EDIT',
			'The title bar and phase headings of a vertical chart Visio built are not editable yet.',
		);
	const { swim, phase } = ensureLists(scope, frame);
	if (edit.style !== undefined) setUserRow(frame, 'visCFFStyle', String(edit.style), '');
	if (edit.rotateLabels !== undefined)
		setUserRow(swim, 'visRotateLabel', flagText(edit.rotateLabels), '');
	if (titleChanges) setUserRow(frame, 'visShowTitle', flagText(edit.showTitle!), '');
	if (phaseChanges) setUserRow(frame, 'visShowPhase', flagText(edit.showSeparators!), '');
	// Cells first, then the boxes: placing a shape recalculates it from the new cells.
	recalculateMasterInstances(scope.root, scope.masters);
	if (titleChanges) {
		const box = boundsOf([frame]);
		const band = headingBand(scope, frame, TITLE_BAND) * (edit.showTitle ? 1 : -1);
		if (box) place(scope, frame, { ...box, maxY: box.maxY + band });
	}
	if (phaseChanges) {
		const box = boundsOf([frame]);
		const lanes = lanesOf(scope, frame, before.direction);
		const top = Math.max(...lanes.map((lane) => boundsOf([lane])?.maxY ?? -Infinity));
		const band = headingBand(scope, phase, PHASE_BAND) * (edit.showSeparators ? 1 : -1);
		if (box && Number.isFinite(top)) relayout(scope, frame, undefined, box.maxY - top + band);
	}
	recalculateMasterInstances(scope.root, scope.masters);
	refreshSettings(scope, frame, swim);
}
