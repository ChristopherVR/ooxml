import type { VisioPackageLimits } from './package-common';
import { fail } from './package-common';
import { indexedPart, visioXml } from './parts';
import { attribute, children } from './sheet';
import { addMasterToDrawing, shapesOf, type EnsuredStencilMaster } from './edit-master-add';
import { stencilMasterContents, stencilMasterEntry } from './stencil-master-part';
import { visioBuiltInMaster } from './stencil-masters';

export type { EnsuredStencilMaster };

/**
 * Make a built-in stencil master part of the drawing's document stencil, as Visio copies a master
 * into a drawing the first time it is dropped. A copy made earlier is found by its UniqueID and
 * reused. So is a master of the drawing with the same universal name (Visio's own "Process" when
 * the drawing came from Visio) as long as it is one droppable 2-D shape: the built-in stencils
 * stand in for Visio's stencil files, so the drawing's real master is the better one to instance.
 */
export async function ensureStencilMaster(
	source: Uint8Array,
	master: string,
	limits: VisioPackageLimits,
	maxOutput: number,
	deadline: number,
	check: () => void,
): Promise<EnsuredStencilMaster> {
	const builtIn = visioBuiltInMaster(master);
	if (!builtIn) fail('EDIT_TARGET_NOT_FOUND', 'The stencil has no such master.');
	const name = builtIn.master.name;
	const same = (value: string | undefined, other: string) =>
		value?.toLowerCase() === other.toLowerCase();
	return addMasterToDrawing(
		source,
		{
			uniqueId: builtIn.uniqueId,
			name,
			async reuse({ pkg, existingPart, entries, check: tick }) {
				for (const entry of entries) {
					tick();
					if (!existingPart || !same(attribute(entry, 'NameU') ?? attribute(entry, 'Name'), name))
						continue;
					try {
						const part = await indexedPart(pkg, existingPart, entry, 'master');
						const roots = shapesOf(await visioXml(pkg, part, 'MasterContents')).filter(
							(shape) => attribute(shape, 'Del') !== '1',
						);
						const cells = new Set(children(roots[0], 'Cell').map((cell) => attribute(cell, 'N')));
						if (roots.length === 1 && !cells.has('BeginX') && !roots[0]!.hasAttribute('Master'))
							return attribute(entry, 'ID') ?? '';
					} catch {
						// A master that cannot be read is not reused; ours is added beside it.
					}
				}
				return undefined;
			},
			entry: (masters, id, unique, relId) =>
				stencilMasterEntry(masters, builtIn, id, unique, relId),
			contents: (document) => stencilMasterContents(document.namespaceURI!, document, builtIn),
		},
		limits,
		maxOutput,
		deadline,
		check,
	);
}
