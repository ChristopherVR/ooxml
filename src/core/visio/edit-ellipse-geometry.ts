import { attribute, children } from './sheet';
import { cells, numeric, setCell } from './edit-geometry-cells';
import { fail } from './package-common';

// Native DrawOval's orthogonal control axes, shared by creation and resize proof.
const controls = [
	['X', 'Width', 0.5],
	['Y', 'Height', 0.5],
	['A', 'Width', 1],
	['B', 'Height', 0.5],
	['C', 'Width', 0.5],
	['D', 'Height', 1],
] as const;
export function appendEllipseGeometry(shape: Element, width: number, height: number): void {
	shape.appendChild(ellipseGeometrySection(shape, width, height));
}
/** A detached native DrawOval Geometry section (IX 0) for a shape of this size. */
export function ellipseGeometrySection(shape: Element, width: number, height: number): Element {
	const node = (name: string) => shape.ownerDocument!.createElementNS(shape.namespaceURI, name);
	const section = node('Section');
	section.setAttribute('N', 'Geometry');
	section.setAttribute('IX', '0');
	for (const name of ['NoFill', 'NoLine', 'NoShow', 'NoSnap', 'NoQuickDrag'])
		setCell(section, name, 0);
	const row = node('Row');
	row.setAttribute('IX', '1');
	row.setAttribute('T', 'Ellipse');
	for (const [name, dimension, factor] of controls) {
		setCell(row, name, (dimension === 'Width' ? width : height) * factor, `${dimension}*${factor}`);
		if (!['X', 'Y'].includes(name)) cells(row).get(name)!.setAttribute('U', 'DL');
	}
	section.appendChild(row);
	return section;
}
/** Admit the native orthogonal row; arbitrary ellipse control axes need their own proof. */
export function assertEllipseResizeRow(shape: Element, row: Element): void {
	const sections = children(shape, 'Section').filter(
		(section) => attribute(section, 'N') === 'Geometry',
	);
	if (
		sections.length !== 1 ||
		attribute(sections[0], 'IX') !== '0' ||
		children(sections[0], 'Row').length !== 1 ||
		attribute(row, 'IX') !== '1'
	)
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'Only a single native ellipse geometry row can be resized.');
	const local = cells(shape),
		coordinates = cells(row);
	if (coordinates.size !== controls.length)
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'Native ellipse control cells are required.');
	for (const [name, dimension, factor] of controls) {
		const cell = coordinates.get(name),
			formula = attribute(cell, 'F')?.replace(/\s/g, '').toUpperCase();
		const expected = numeric(local.get(dimension)) * factor;
		if (
			formula !== `${dimension}*${factor}`.toUpperCase() ||
			Math.abs(numeric(cell) - expected) > 1e-9 * Math.max(1, Math.abs(expected))
		)
			fail(
				'UNSUPPORTED_GEOMETRY_EDIT',
				'Native dimension-dependent ellipse control formulas and caches are required.',
			);
	}
}
