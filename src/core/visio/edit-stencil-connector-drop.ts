import { visioWalkPath } from './connector-route-walk';
import { cells, setCell } from './edit-geometry-cells';
import type { VisioMasterInstanceEdit } from './edit-master-instance';
import { isDynamicConnectorMaster, registerStencilShape } from './edit-stencil-connector';
import { writeStencilConnector } from './edit-stencil-connector-write';
import { visioFormulaCachedValue } from './formula';
import { fail, VisioPackageError } from './package-common';
import { attribute, children } from './sheet';

const number = (shape: Element, name: string): number => {
	const node = children(shape, 'Cell').find((cell) => attribute(cell, 'N') === name);
	const value =
		node && !node.hasAttribute('E')
			? visioFormulaCachedValue(attribute(node, 'V') ?? '', attribute(node, 'U')).value
			: NaN;
	if (!Number.isFinite(value) || Math.abs(value) > 1e6)
		fail('UNSUPPORTED_MASTER_INSTANCE', `The connector master has no usable ${name}.`);
	return value;
};

/**
 * Fill a new instance of a connector master (Visio's Dynamic connector) dropped on a page, as
 * Visio saves one: the two free ends, the transform that follows them, the one-bend path and the
 * triggers Visio gives an unglued connector (they name the connector's own sheet). The ends are
 * the edit's, or the master's own ends carried to the drop point.
 */
export function dropStencilConnector(
	shape: Element,
	master: Element,
	edit: VisioMasterInstanceEdit,
): void {
	const dx = edit.x - number(master, 'PinX'),
		dy = edit.y - number(master, 'PinY');
	const translatedBegin = {
		x: number(master, 'BeginX') + dx,
		y: number(master, 'BeginY') + dy,
	};
	const translatedEnd = { x: number(master, 'EndX') + dx, y: number(master, 'EndY') + dy };
	const begin = edit.begin ?? translatedBegin;
	const end = edit.end ?? translatedEnd;
	if ([begin.x, begin.y, end.x, end.y].some((value) => Math.abs(value) > 1e6))
		fail('INVALID_EDIT', 'Dropped connector coordinates exceed limits.');
	if (Math.hypot(end.x - begin.x, end.y - begin.y) < 1e-6)
		fail('INVALID_EDIT', 'The ends of a connector cannot be in the same place.');
	if (!isDynamicConnectorMaster(master) || attribute(master, 'Type') === 'Group') {
		if (
			Math.hypot(begin.x - translatedBegin.x, begin.y - translatedBegin.y) > 1e-9 ||
			Math.hypot(end.x - translatedEnd.x, end.y - translatedEnd.y) > 1e-9
		)
			fail(
				'UNSUPPORTED_MASTER_INSTANCE',
				'This connector master must keep its original end-to-end geometry on drop.',
			);
		const source = cells(master);
		setCell(shape, 'PinX', edit.x, source.get('PinX')?.hasAttribute('F') ? 'Inh' : undefined);
		setCell(shape, 'PinY', edit.y, source.get('PinY')?.hasAttribute('F') ? 'Inh' : undefined);
		setCell(shape, 'BeginX', begin.x);
		setCell(shape, 'BeginY', begin.y);
		setCell(shape, 'EndX', end.x);
		setCell(shape, 'EndY', end.y);
		return;
	}
	registerStencilShape(shape, master);
	try {
		writeStencilConnector(
			shape,
			'right-angle',
			visioWalkPath({ point: begin }, { point: end }) ?? [begin, end],
		);
	} catch (error) {
		if (error instanceof VisioPackageError && error.code === 'UNSUPPORTED_GEOMETRY_EDIT')
			fail(
				'UNSUPPORTED_MASTER_INSTANCE',
				'The Dynamic connector master cannot be laid out with these ends.',
			);
		throw error;
	}
	const trigger = `_XFTRIGGER(Sheet.${attribute(shape, 'ID')}!EventXFMod)`;
	// Visio's cell order: the triggers sit before the text pin.
	const text = children(shape, 'Cell').find((cell) => attribute(cell, 'N') === 'TxtPinX');
	for (const name of ['BegTrigger', 'EndTrigger']) {
		setCell(shape, name, 1, trigger);
		const node = children(shape, 'Cell').find((cell) => attribute(cell, 'N') === name)!;
		if (text) shape.insertBefore(node, text);
	}
}
