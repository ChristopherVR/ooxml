import { attribute, children } from './sheet';
import { fail } from './package-common';
import { executableCellFormula } from './cell-formula';

/** Native Visio stores Title and AlternativeText in these named User rows. */
export function writeShapeAltText(
	shape: Element,
	title: string,
	description: string,
	inherited = false,
): boolean {
	const sections = children(shape, 'Section').filter(
		(section) => attribute(section, 'N') === 'User',
	);
	if (sections.length > 1 || sections[0]?.hasAttribute('Del'))
		fail('UNSUPPORTED_METADATA_EDIT', 'Ambiguous or deleted User sections cannot be edited.');
	let section = sections[0];
	let changed = false;
	for (const [name, value] of [
		['visAltText', title],
		['visAltDescription', description],
	] as const) {
		const rows = children(section, 'Row').filter((row) => attribute(row, 'N') === name);
		if (rows.length > 1 || rows[0]?.hasAttribute('Del'))
			fail(
				'UNSUPPORTED_METADATA_EDIT',
				'Ambiguous or deleted accessibility rows cannot be edited.',
			);
		let row = rows[0];
		const cells = children(row, 'Cell').filter((cell) => attribute(cell, 'N') === 'Value');
		if (
			cells.length > 1 ||
			cells[0]?.hasAttribute('E') ||
			executableCellFormula(attribute(cells[0], 'F'))
		)
			fail('UNSUPPORTED_METADATA_EDIT', 'Formula-driven alternative text cannot be replaced.');
		let cell = cells[0];
		if (attribute(cell, 'V') === value || (!row && !value && !inherited)) continue;
		const create = (tag: string) => shape.ownerDocument!.createElementNS(shape.namespaceURI, tag);
		if (!section) {
			section = create('Section');
			section.setAttribute('N', 'User');
			const after = Array.from(shape.childNodes).find(
				(node) =>
					node.nodeType === 1 &&
					['Text', 'Data1', 'Data2', 'Data3', 'ForeignData', 'Shapes'].includes(
						(node as Element).localName,
					),
			);
			shape.insertBefore(section, after ?? null);
		}
		if (!row) {
			row = create('Row');
			row.setAttribute('N', name);
			section.appendChild(row);
		}
		if (!cell) {
			cell = create('Cell');
			cell.setAttribute('N', 'Value');
			row.insertBefore(cell, row.firstChild);
		}
		cell.setAttribute('V', value);
		cell.setAttribute('U', 'STR');
		cell.removeAttribute('F');
		changed = true;
	}
	return changed;
}
