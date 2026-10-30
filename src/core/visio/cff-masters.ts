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

/**
 * The UniqueIDs of the masters this package writes for the charts it creates (`cff-author.ts`).
 * They carry Visio's names and structure rows but not its formulas, so the editor treats their
 * instances as it treats its own plain shapes (direct cells), and only tells Visio what they are.
 */
export const CFF_AUTHORED_IDS: Record<CffRole, string> = {
	frame: '{0F0C7A10-0001-0000-8E40-00608CF305B2}',
	'swimlane-list': '{0F0C7A10-0002-0000-8E40-00608CF305B2}',
	'phase-list': '{0F0C7A10-0003-0000-8E40-00608CF305B2}',
	lane: '{0F0C7A10-0004-0000-8E40-00608CF305B2}',
	separator: '{0F0C7A10-0005-0000-8E40-00608CF305B2}',
};
const AUTHORED_ROLES = new Map(
	(Object.entries(CFF_AUTHORED_IDS) as [CffRole, string][]).map(([role, id]) => [id, role]),
);

/** The drawing's masters of the cross-functional flowchart, with their top-level shapes. */
export interface CffMasterSet extends CffMasters {
	/** The role of a master Visio built; masters this package wrote answer nothing. */
	roleOf(masterId: string | null | undefined): CffRole | undefined;
	/** The ID of the master this package wrote for a role, when the drawing has it. */
	authoredId(role: CffRole): string | undefined;
	/** The part of that master, which a page that uses it must be related to. */
	authoredPart(role: CffRole): string | undefined;
	readonly empty: boolean;
}
export const noCffMasters: CffMasterSet = {
	get: () => undefined,
	roleOf: () => undefined,
	authoredId: () => undefined,
	authoredPart: () => undefined,
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
	const authored = new Map<CffRole, { id: string; part: string }>();
	for (const master of children(list, 'Master')) {
		const id = attribute(master, 'ID');
		const role =
			ROLES[(attribute(master, 'NameU') ?? attribute(master, 'Name') ?? '').toLowerCase()];
		if (!id) continue;
		const own = AUTHORED_ROLES.get((attribute(master, 'UniqueID') ?? '').toUpperCase());
		if (!role && !own) continue;
		const part = await indexedPart(pkg, mastersPart, master, 'master');
		if (own) {
			authored.set(own, { id, part });
			continue;
		}
		const root = await visioXml(pkg, part, 'MasterContents');
		const top = children(children(root, 'Shapes')[0], 'Shape');
		if (top.length !== 1) continue;
		shapes.set(id, top[0]!);
		roles.set(id, role!);
	}
	if (!shapes.size && !authored.size) return noCffMasters;
	return {
		get: (id) => shapes.get(id),
		roleOf: (id) => (id ? roles.get(id) : undefined),
		authoredId: (role) => authored.get(role)?.id,
		authoredPart: (role) => authored.get(role)?.part,
		empty: !shapes.size,
	};
}
