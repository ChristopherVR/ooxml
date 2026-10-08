import type { VisioPackage } from './package';
import { assertShapeOrderPackageScope } from './edit-shape-order';
import { attribute } from './sheet';
import { executableCellFormula } from './cell-formula';
import { analyzeVisioFormula } from './formula';
import { fail } from './package-common';

/** Adding shapes can affect container lookups and resolve formerly missing references. */
export async function assertDuplicateScope(
	pkg: VisioPackage,
	pagePaths: ReadonlySet<string>,
	targetPage: Element,
	newIds: ReadonlySet<string>,
	check: () => void,
): Promise<void> {
	await assertShapeOrderPackageScope(pkg, check);
	const inspect = (root: Element) => {
		for (const node of [root, ...Array.from(root.getElementsByTagName('*'))]) {
			check();
			const source = executableCellFormula(attribute(node, 'F'));
			if (!source) continue;
			if (
				analyzeVisioFormula(source).references.some(
					(ref) => ref.shapeId !== undefined && newIds.has(ref.shapeId),
				)
			)
				fail('EDIT_DUPLICATE_DEPENDENCY', 'New IDs are already referenced by existing formulas.');
		}
	};
	// Sheet IDs are page-local. Independent pages can already contain these IDs and references.
	inspect(targetPage);
	for (const path of pkg.paths())
		if (/^visio\/.*\.xml$/i.test(path) && !pagePaths.has(path)) inspect(await pkg.readXml(path));
}
