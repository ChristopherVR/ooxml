import { attribute, children } from './sheet';
import { fail } from './package-common';
import { executableCellFormula } from './cell-formula';
import type { VisioMetadataEdit, VisioHyperlinkFields } from './edit-metadata-commands';

const HYPERLINK_CELLS: readonly [string, string][] = [
	['Description', ''],
	['Address', ''],
	['SubAddress', ''],
	['ExtraInfo', ''],
	['Frame', ''],
	['SortKey', ''],
	['NewWindow', '0'],
	['Default', '0'],
	['Invisible', '0'],
];

/** One unique local page shape, never a master instance whose sections would merge. */
function metadataShape(root: Element, shapeId: string): Element {
	const matches = Array.from(root.getElementsByTagNameNS(root.namespaceURI, 'Shape')).filter(
		(shape) => attribute(shape, 'ID') === shapeId,
	);
	if (matches.length !== 1) fail('EDIT_TARGET_NOT_FOUND', 'A unique page shape is required.');
	const shape = matches[0]!;
	for (let node: Element | null = shape; node && node !== root; node = node.parentElement)
		if (
			node.localName === 'Shape' &&
			['Master', 'MasterShape', 'Del'].some((name) => node!.hasAttribute(name))
		)
			fail(
				'UNSUPPORTED_METADATA_EDIT',
				'Links and ScreenTips on master instances need master section inheritance, which is not supported.',
			);
	return shape;
}

/** Refuse when any page formula could read the cells being changed. */
function assertUnreferenced(roots: ReadonlyMap<string, Element>, check: () => void): void {
	for (const root of roots.values())
		for (const cell of Array.from(root.getElementsByTagNameNS(root.namespaceURI, 'Cell'))) {
			check();
			const formula = executableCellFormula(attribute(cell, 'F'));
			if (formula && /\b(?:Comment|Hyperlinks?)\b/i.test(formula))
				fail(
					'UNSUPPORTED_METADATA_EDIT',
					'A ShapeSheet formula may read ScreenTip or hyperlink cells; recalculation is not supported.',
				);
		}
}

function writableCell(owner: Element, name: string): Element | undefined {
	const matches = children(owner, 'Cell').filter((cell) => attribute(cell, 'N') === name);
	if (matches.length > 1) fail('UNSUPPORTED_METADATA_EDIT', `Duplicate ${name} cells.`);
	const cell = matches[0];
	if (cell && (cell.hasAttribute('E') || executableCellFormula(attribute(cell, 'F'))))
		fail('UNSUPPORTED_METADATA_EDIT', `The ${name} cell is formula driven and cannot be replaced.`);
	return cell;
}
function setText(owner: Element, name: string, value: string): boolean {
	let cell = writableCell(owner, name);
	if (cell && attribute(cell, 'V') === value) return false;
	if (!cell) {
		cell = owner.ownerDocument!.createElementNS(owner.namespaceURI, 'Cell');
		cell.setAttribute('N', name);
		const before = Array.from(owner.childNodes).find(
			(node) => node.nodeType === 1 && (node as Element).localName !== 'Cell',
		);
		owner.insertBefore(cell, before ?? null);
	}
	cell.setAttribute('V', value);
	cell.removeAttribute('F');
	cell.removeAttribute('U');
	return true;
}

function hyperlinkSection(shape: Element, create: boolean): Element | undefined {
	const sections = children(shape, 'Section').filter(
		(node) => attribute(node, 'N') === 'Hyperlink',
	);
	if (sections.length > 1) fail('UNSUPPORTED_METADATA_EDIT', 'Duplicate Hyperlink sections.');
	if (sections[0]?.hasAttribute('Del'))
		fail('UNSUPPORTED_METADATA_EDIT', 'Deleted Hyperlink sections are not supported.');
	if (sections[0] || !create) return sections[0];
	const section = shape.ownerDocument!.createElementNS(shape.namespaceURI, 'Section');
	section.setAttribute('N', 'Hyperlink');
	const after = Array.from(shape.childNodes).find(
		(node) =>
			node.nodeType === 1 &&
			['Text', 'Data1', 'Data2', 'Data3', 'ForeignData', 'Shapes'].includes(
				(node as Element).localName,
			),
	);
	shape.insertBefore(section, after ?? null);
	return section;
}

function writeHyperlink(
	section: Element,
	row: string | undefined,
	link: VisioHyperlinkFields,
): boolean {
	const rows = children(section, 'Row');
	let target: Element | undefined;
	if (row !== undefined) {
		const matches = rows.filter((node) => attribute(node, 'N') === row);
		if (matches.length !== 1 || matches[0]!.hasAttribute('Del'))
			fail('EDIT_TARGET_NOT_FOUND', 'The hyperlink row does not exist.');
		target = matches[0]!;
	} else {
		const names = new Set(rows.map((node) => attribute(node, 'N')));
		let index = 1;
		while (names.has(`Row_${index}`)) index++;
		target = section.ownerDocument!.createElementNS(section.namespaceURI, 'Row');
		target.setAttribute('N', `Row_${index}`);
		for (const [name, value] of HYPERLINK_CELLS) {
			const cell = section.ownerDocument!.createElementNS(section.namespaceURI, 'Cell');
			cell.setAttribute('N', name);
			cell.setAttribute('V', value);
			target.appendChild(cell);
		}
		section.appendChild(target);
	}
	let changed = row === undefined;
	changed = setText(target, 'Description', link.description) || changed;
	changed = setText(target, 'Address', link.address) || changed;
	changed = setText(target, 'SubAddress', link.subAddress) || changed;
	return changed;
}

/** Apply a hyperlink or ScreenTip edit to cloned page roots. Returns whether the page changed. */
export function applyMetadataEdit(
	roots: ReadonlyMap<string, Element>,
	edit: VisioMetadataEdit,
	check: () => void,
): boolean {
	const root = roots.get(edit.pageId);
	if (!root) fail('EDIT_TARGET_NOT_FOUND', 'Page does not exist.');
	const shape = metadataShape(root, edit.shapeId);
	assertUnreferenced(roots, check);
	if (edit.type === 'set-shape-screentip') {
		if (edit.text) return setText(shape, 'Comment', edit.text);
		const cell = writableCell(shape, 'Comment');
		if (!cell) return false;
		shape.removeChild(cell);
		return true;
	}
	if (edit.hyperlink)
		return writeHyperlink(hyperlinkSection(shape, true)!, edit.row, edit.hyperlink);
	const section = hyperlinkSection(shape, false);
	const row = children(section, 'Row').filter((node) => attribute(node, 'N') === edit.row);
	if (!section || row.length !== 1)
		fail('EDIT_TARGET_NOT_FOUND', 'The hyperlink row does not exist.');
	for (const cell of children(row[0], 'Cell'))
		if (cell.hasAttribute('E') || executableCellFormula(attribute(cell, 'F')))
			fail('UNSUPPORTED_METADATA_EDIT', 'Formula-driven hyperlinks cannot be removed.');
	section!.removeChild(row[0]!);
	if (!children(section, 'Row').length) shape.removeChild(section!);
	return true;
}
