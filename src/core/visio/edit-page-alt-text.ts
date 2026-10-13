import type { VisioPackage } from './package';
import type { VisioPageAltTextEdit } from './edit-page-setup-commands';
import { pageById } from './edit-page-setup';
import { children, attribute } from './sheet';
import { fail } from './package-common';
import { executableCellFormula } from './cell-formula';
import { writeShapeAltText } from './edit-alt-text';

export async function setVisioPageAltText(
	pkg: VisioPackage,
	pagesPart: string,
	pages: Element,
	dirty: Map<string, Element>,
	command: VisioPageAltTextEdit,
	check: () => void,
): Promise<void> {
	const page = pageById(pages, command.pageId);
	const sheets = children(page, 'PageSheet');
	if (sheets.length > 1) fail('UNSUPPORTED_METADATA_EDIT', 'Ambiguous PageSheet metadata.');
	for (const path of pkg.paths()) {
		if (!/^visio\/.*\.xml$/i.test(path)) continue;
		const part = path === pagesPart ? pages : (dirty.get(path) ?? (await pkg.readXml(path)));
		for (const cell of Array.from(part.getElementsByTagNameNS(part.namespaceURI, 'Cell'))) {
			check();
			const formula = executableCellFormula(attribute(cell, 'F'));
			if (formula && /\bUser\s*\.\s*visAlt(?:Text|Description)\b/i.test(formula))
				fail(
					'UNSUPPORTED_METADATA_EDIT',
					'A formula reads alternative text; dependent caches cannot be recalculated.',
				);
		}
	}
	let sheet = sheets[0];
	if (!sheet) {
		if (!command.title && !command.description) return;
		sheet = page.ownerDocument!.createElementNS(page.namespaceURI, 'PageSheet');
		page.insertBefore(sheet, page.firstChild);
	}
	if (writeShapeAltText(sheet, command.title, command.description)) dirty.set(pagesPart, pages);
}
