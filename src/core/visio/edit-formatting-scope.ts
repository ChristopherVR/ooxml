import type { VisioPackage } from './package';
import { attribute, VISIO_NS, VISIO_LEGACY_NS } from './sheet';
import { executableCellFormula, inertDoubleClickFormula } from './cell-formula';
import { analyzeVisioFormula } from './formula';
import { indexCells } from './edit-recalculate-index';
import { fail } from './package-common';

/** Nonnumeric font/color caches cannot be recalculated by the numeric engine.
 * Refuse every dependent formula, including metadata/styles/masters, before writes.
 */
export async function assertFormattingDependencies(
	pkg: VisioPackage,
	pagePaths: ReadonlySet<string>,
	roots: ReadonlyMap<string, Element>,
	pageId: string,
	shape: Element,
	changed: ReadonlyMap<string, Element | undefined>,
	check: () => void,
): Promise<void> {
	indexCells(roots, { check });
	const shapeId = attribute(shape, 'ID')!;
	const changedNames = new Set([...changed.keys()].map((name) => name.toLowerCase()));
	// Native ShapeSheet names differ from XML section/row names for text cells.
	// https://learn.microsoft.com/en-us/office/client-developer/visio/style-cell-character-section
	const canonical = (name: string) =>
		name
			.toLowerCase()
			.replace(/^char\.([a-z]+)$/, 'character.0.$1')
			.replace(/^para\.([a-z]+)$/, 'paragraph.0.$1');
	const replaced = new Set(changed.values());
	const inspect = (root: Element, sourcePageId?: string) => {
		for (const node of [root, ...Array.from(root.getElementsByTagName('*'))]) {
			check();
			if (![VISIO_NS, VISIO_LEGACY_NS].includes(node.namespaceURI ?? '') || replaced.has(node))
				continue;
			const source = executableCellFormula(attribute(node, 'F'));
			if (!source || inertDoubleClickFormula(attribute(node, 'N') ?? '', source)) continue;
			let localShapeId: string | undefined;
			let parent: Node | null = node;
			while (parent?.nodeType === 1) {
				if ((parent as Element).localName === 'Shape') {
					localShapeId = attribute(parent as Element, 'ID');
					break;
				}
				parent = parent.parentNode;
			}
			const analysis = analyzeVisioFormula(source);
			if (
				analysis.dynamic ||
				analysis.references.some(
					(ref) =>
						changedNames.has(canonical(ref.cell)) &&
						(sourcePageId === undefined || sourcePageId === pageId) &&
						(ref.shapeId === shapeId ||
							(ref.shapeId === undefined &&
								(localShapeId === shapeId || sourcePageId === undefined))),
				)
			)
				fail(
					'EDIT_UNSUPPORTED_FORMAT_DEPENDENCY',
					'Formatting has affected or unknown formula dependencies; dependent caches cannot be recalculated safely.',
				);
		}
	};
	for (const [id, root] of roots) inspect(root, id);
	for (const path of pkg.paths()) {
		if (!/^visio\/.*\.xml$/i.test(path) || pagePaths.has(path)) continue;
		inspect(await pkg.readXml(path));
	}
}
