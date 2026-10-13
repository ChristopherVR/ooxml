import type { VisioPackage } from './package';
import type { VisioShapeAltTextEdit } from './edit-metadata-commands';
import { attribute, children } from './sheet';
import { fail } from './package-common';
import { executableCellFormula } from './cell-formula';
import { textTarget } from './edit-text-target';
import { masterTemplate } from './edit-text-scope';
import { writeShapeAltText } from './edit-alt-text';

/** Override only the native accessibility rows; geometry and unrelated master data stay inherited. */
export async function applyShapeAltText(
	pkg: VisioPackage,
	pagePaths: ReadonlySet<string>,
	roots: ReadonlyMap<string, Element>,
	edit: VisioShapeAltTextEdit,
	check: () => void,
): Promise<boolean> {
	const root = roots.get(edit.pageId);
	if (!root) fail('EDIT_TARGET_NOT_FOUND', 'Page does not exist.');
	const target = textTarget(root, edit.shapeId, check);
	const inspect = (part: Element) => {
		for (const cell of Array.from(part.getElementsByTagNameNS(part.namespaceURI, 'Cell'))) {
			check();
			const formula = executableCellFormula(attribute(cell, 'F'));
			if (formula && /\bUser\s*\.\s*visAlt(?:Text|Description)\b/i.test(formula))
				fail(
					'UNSUPPORTED_METADATA_EDIT',
					'A formula reads alternative text; dependent caches cannot be recalculated.',
				);
		}
	};
	for (const page of roots.values()) inspect(page);
	for (const path of pkg.paths()) {
		if (/^visio\/.*\.xml$/i.test(path) && !pagePaths.has(path)) inspect(await pkg.readXml(path));
	}
	if (target.masterId !== undefined) {
		const template = await masterTemplate(pkg)(target.masterId, target.masterShapeId);
		if (template.hasAttribute('Master') || template.hasAttribute('MasterShape'))
			fail('UNSUPPORTED_METADATA_EDIT', 'Chained master accessibility metadata is not supported.');
		// Validate inherited metadata before creating a local override, including formula-driven values.
		// A literal local override can be changed without replacing the master's own formula.
		const inherited = template.cloneNode(true) as Element;
		for (const section of children(inherited, 'Section').filter(
			(section) => attribute(section, 'N') === 'User',
		)) {
			for (const row of children(section, 'Row')) {
				const localRow = children(target.node, 'Section')
					.filter((s) => attribute(s, 'N') === 'User')
					.flatMap((s) => children(s, 'Row'))
					.find((r) => attribute(r, 'N') === attribute(row, 'N'));
				const local = children(localRow, 'Cell').find((cell) => attribute(cell, 'N') === 'Value');
				if (local && attribute(local, 'F') !== 'Inh') {
					for (const cell of children(row, 'Cell').filter(
						(cell) => attribute(cell, 'N') === 'Value',
					))
						row.removeChild(cell);
					row.appendChild(local.cloneNode(true));
				}
			}
		}
		writeShapeAltText(inherited, edit.title, edit.description, true);
	}
	return writeShapeAltText(
		target.node,
		edit.title,
		edit.description,
		target.masterId !== undefined,
	);
}
