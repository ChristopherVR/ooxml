import type { VisioPackage } from './package';
import { indexedPart, related, visioXml } from './parts';
import { attribute, children } from './sheet';
import type { CffMasters } from './cff-recalc';

/** What a master of Visio's cross-functional flowchart is, by its name. */
export type CffRole = 'frame' | 'lane' | 'separator' | 'swimlane-list' | 'phase-list';

const ROLES: Record<string, CffRole> = {
	'cff container': 'frame',
	swimlane: 'lane',
	separator: 'separator',
	'swimlane list': 'swimlane-list',
	'phase list': 'phase-list',
};

/** The drawing's masters of the cross-functional flowchart, with their top-level shapes. */
export interface CffMasterSet extends CffMasters {
	roleOf(masterId: string | null | undefined): CffRole | undefined;
	readonly empty: boolean;
}
export const noCffMasters: CffMasterSet = {
	get: () => undefined,
	roleOf: () => undefined,
	empty: true,
};

/**
 * Read the masters Visio's Cross-Functional Flowchart template puts in a drawing (CFF Container,
 * Swimlane List, Phase List, Swimlane, Separator, found by name) so the lane commands can
 * recognise a chart Visio built and recalculate its instances.
 */
export async function loadCffMasters(pkg: VisioPackage): Promise<CffMasterSet> {
	const documentPart = await related(pkg, '', 'document');
	const mastersPart = documentPart && (await related(pkg, documentPart, 'masters', false));
	if (!mastersPart) return noCffMasters;
	const list = await visioXml(pkg, mastersPart, 'Masters');
	const shapes = new Map<string, Element>();
	const roles = new Map<string, CffRole>();
	for (const master of children(list, 'Master')) {
		const id = attribute(master, 'ID');
		const role =
			ROLES[(attribute(master, 'NameU') ?? attribute(master, 'Name') ?? '').toLowerCase()];
		if (!id || !role) continue;
		const root = await visioXml(
			pkg,
			await indexedPart(pkg, mastersPart, master, 'master'),
			'MasterContents',
		);
		const top = children(children(root, 'Shapes')[0], 'Shape');
		if (top.length !== 1) continue;
		shapes.set(id, top[0]!);
		roles.set(id, role);
	}
	if (!shapes.size) return noCffMasters;
	return {
		get: (id) => shapes.get(id),
		roleOf: (id) => (id ? roles.get(id) : undefined),
		empty: false,
	};
}
