import { NS, parseXml } from '../xml/index';
import { relationshipsPartFor, resolvePartPath } from '../opc/relationships';
import { attribute } from './sheet';
import { copy } from './edit-master-add';
import { shapeIdOf, topLevelShapes, userValue } from './edit-container-members';
import {
	VISIO_SEPARATOR_ROW,
	VISIO_SWIMLANE_FRAME_ROW,
	VISIO_SWIMLANE_LANE_ROW,
} from './shape-structure';
import { CFF_AUTHORED_IDS, type CffMasterSet, type CffRole } from './cff-masters';

const MASTER_RELATIONSHIP = 'http://schemas.microsoft.com/visio/2010/relationships/master';
const CFF_ROLES = Object.keys(CFF_AUTHORED_IDS) as CffRole[];

/** Which part of a cross-functional flowchart a shape this package wrote is, by its own rows. */
function partOf(shape: Element): CffRole | undefined {
	if (userValue(shape, VISIO_SWIMLANE_FRAME_ROW) === '1') return 'frame';
	if (userValue(shape, VISIO_SWIMLANE_LANE_ROW) === '1') return 'lane';
	if (userValue(shape, VISIO_SEPARATOR_ROW) === '1') return 'separator';
	if (userValue(shape, 'msvStructureType')?.toLowerCase() !== 'list') return undefined;
	const kind = userValue(shape, 'msvSDListRequiredCategories');
	return kind === 'Swimlane' ? 'swimlane-list' : kind === 'Phase' ? 'phase-list' : undefined;
}

/** Take the masters this package wrote off the shapes that name them (see `stampCffMasters`). */
export function unstampCffMasters(root: Element, masters: CffMasterSet): void {
	const ids = new Set(CFF_ROLES.map((role) => masters.authoredId(role)).filter(Boolean));
	if (!ids.size) return;
	for (const shape of topLevelShapes(root))
		if (ids.has(attribute(shape, 'Master'))) shape.removeAttribute('Master');
}

/**
 * Make the frame, lists, lanes and separators of the charts this package builds instances of the
 * hidden flowchart masters the drawing holds (`cff-author.ts`), as Visio's own template does.
 * The shapes keep every cell they carry, so nothing about them changes but the master they
 * name. A shape that already names a master (a chart Visio built) is left alone.
 */
export function stampCffMasters(root: Element, masters: CffMasterSet): Set<CffRole> {
	const used = new Set<CffRole>();
	for (const shape of topLevelShapes(root)) {
		const id = attribute(shape, 'Master');
		const role = CFF_ROLES.find((item) => masters.authoredId(item) === id);
		if (role && id) used.add(role);
	}
	for (const shape of topLevelShapes(root)) {
		if (attribute(shape, 'Master') || !shapeIdOf(shape)) continue;
		const role = partOf(shape);
		const id = role && masters.authoredId(role);
		if (!id) continue;
		shape.setAttribute('Master', id);
		used.add(role!);
	}
	return used;
}

/**
 * A page names the masters it uses in its relationships; Visio does not load an instance whose
 * master the page is not related to. Returns the relationships part to save, or nothing when the
 * page already relates to every master in `roles`.
 */
export async function relateCffMasters(
	pkg: { readXml(part: string, root: string): Promise<Element> },
	parts: ReadonlyMap<string, unknown>,
	pagePath: string,
	roles: ReadonlySet<CffRole>,
	masters: CffMasterSet,
): Promise<{ part: string; rels: Element } | undefined> {
	const part = relationshipsPartFor(pagePath);
	const rels = parts.has(part)
		? copy(await pkg.readXml(part, 'Relationships'))
		: parseXml(`<Relationships xmlns="${NS.rels}"/>`).documentElement;
	const existing = () => Array.from(rels.getElementsByTagNameNS(rels.namespaceURI, 'Relationship'));
	let changed = false;
	for (const role of roles) {
		const target = masters.authoredPart(role);
		if (!target) continue;
		const linked = existing().some(
			(rel) =>
				rel.getAttribute('Type') === MASTER_RELATIONSHIP &&
				rel.getAttribute('TargetMode') !== 'External' &&
				resolvePartPath(pagePath, rel.getAttribute('Target') ?? '') === target,
		);
		if (linked) continue;
		const ids = new Set(existing().map((rel) => rel.getAttribute('Id') ?? ''));
		let index = 1;
		while (ids.has(`rId${index}`)) index++;
		const relationship = rels.ownerDocument!.createElementNS(rels.namespaceURI, 'Relationship');
		relationship.setAttribute('Id', `rId${index}`);
		relationship.setAttribute('Type', MASTER_RELATIONSHIP);
		// Page and master parts are siblings under visio/ in every package Visio writes.
		const from = pagePath.split('/').slice(0, -1);
		const to = target.split('/');
		let shared = 0;
		while (shared < from.length && shared < to.length - 1 && from[shared] === to[shared]) shared++;
		relationship.setAttribute(
			'Target',
			[...from.slice(shared).map(() => '..'), ...to.slice(shared)].join('/'),
		);
		rels.appendChild(relationship);
		changed = true;
	}
	return changed ? { part, rels } : undefined;
}
