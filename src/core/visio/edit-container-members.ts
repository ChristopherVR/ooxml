import { attribute, children } from './sheet';
import { fail } from './package-common';
import { cells, setCell } from './edit-geometry-cells';
import { containerMemberIds } from './shape-structure';
import {
	VISIO_RELATIONSHIP_LIST_MEMBERS,
	VISIO_RELATIONSHIP_LIST_MEMBER_OF,
	VISIO_RELATIONSHIP_MEMBERS,
	VISIO_RELATIONSHIP_MEMBER_OF,
	dependsOn,
	relationTerms,
	withRelation,
	withoutSheets,
} from './container-relationships';
import {
	VISIO_CONTAINER_MEMBERS_ROW,
	type VisioSetContainerMembersEdit,
} from './edit-diagram-parts-commands';

/** The top-level local shapes of a page (the scope containers operate in). */
export function topLevelShapes(root: Element): Element[] {
	const containers = children(root, 'Shapes');
	return containers.length === 1 ? children(containers[0]!, 'Shape') : [];
}
export const shapeIdOf = (shape: Element): string => attribute(shape, 'ID') ?? '';

function userRow(shape: Element, name: string): Element | undefined {
	for (const section of children(shape, 'Section'))
		if (attribute(section, 'N') === 'User')
			for (const row of children(section, 'Row')) if (attribute(row, 'N') === name) return row;
	return undefined;
}
/** The text of a User row's Value cell, quotes stripped. */
export function userValue(shape: Element, name: string): string | undefined {
	const row = userRow(shape, name);
	const cell = row && children(row, 'Cell').find((node) => attribute(node, 'N') === 'Value');
	return cell ? (attribute(cell, 'V') ?? '').replace(/^"(.*)"$/s, '$1').trim() : undefined;
}
export const isContainerSheet = (shape: Element): boolean =>
	userValue(shape, 'msvStructureType')?.toLowerCase() === 'container';
/** A container's member IDs, from the editor's older private row or its Relationships text. */
export function containerMembersOf(shape: Element): string[] {
	const own = userRow(shape, VISIO_CONTAINER_MEMBERS_ROW);
	return containerMemberIds(
		own ? (userValue(shape, VISIO_CONTAINER_MEMBERS_ROW) ?? '') : undefined,
		attribute(cells(shape).get('Relationships'), 'F'),
	);
}

function writeRelationships(shape: Element, formula: string | undefined): void {
	const node = cells(shape).get('Relationships');
	if (formula) setCell(shape, 'Relationships', 0, formula);
	else if (node) node.parentNode!.removeChild(node);
}
/**
 * Replace the ids of one relationship kind in a shape's Relationships cell, keeping the other
 * kinds. A formula that holds more than relationships is left alone. A container or list keeps
 * its (possibly empty) members term, as Visio writes it.
 */
export function setRelation(shape: Element, kind: number, ids: readonly string[]): void {
	const formula = attribute(cells(shape).get('Relationships'), 'F');
	const keepEmpty = kind === VISIO_RELATIONSHIP_MEMBERS && isContainerSheet(shape);
	const next = withRelation(formula, kind, ids, keepEmpty);
	if (next !== null) writeRelationships(shape, next);
}
/** A member's back references: the containers it belongs to, kept as one DEPENDSON(4, ...). */
function setMembership(member: Element, containerId: string, belongs: boolean): void {
	const formula = attribute(cells(member).get('Relationships'), 'F');
	const ids = dependsOn(formula, VISIO_RELATIONSHIP_MEMBER_OF).filter((id) => id !== containerId);
	if (belongs) ids.push(containerId);
	setRelation(member, VISIO_RELATIONSHIP_MEMBER_OF, ids);
}

/**
 * Record a container's membership in Visio's form: the container's Relationships cell lists its
 * members, each member points back, and the editor's older private User row is dropped.
 */
export function writeContainerMembers(
	root: Element,
	container: Element,
	memberIds: readonly string[],
): void {
	const id = shapeIdOf(container);
	const shapes = new Map(topLevelShapes(root).map((shape) => [shapeIdOf(shape), shape]));
	const previous = containerMembersOf(container);
	for (const memberId of previous)
		if (!memberIds.includes(memberId) && shapes.has(memberId))
			setMembership(shapes.get(memberId)!, id, false);
	for (const memberId of memberIds) {
		const member = shapes.get(memberId);
		if (member) setMembership(member, id, true);
	}
	// Before 2026-10-30 the members were kind 4 and there was no kind 1 term: drop that reading.
	const formula = attribute(cells(container).get('Relationships'), 'F');
	const terms = relationTerms(formula);
	if (terms && !terms.some((term) => term.kind === VISIO_RELATIONSHIP_MEMBERS))
		writeRelationships(
			container,
			withRelation(formula, VISIO_RELATIONSHIP_MEMBER_OF, []) ?? undefined,
		);
	setRelation(container, VISIO_RELATIONSHIP_MEMBERS, memberIds);
	const legacy = userRow(container, VISIO_CONTAINER_MEMBERS_ROW);
	legacy?.parentNode?.removeChild(legacy);
}

/** A list's members in order (the Swimlane List and Phase List of a cross-functional flowchart). */
export function listMembersOf(shape: Element): string[] {
	return [
		...new Set(
			dependsOn(attribute(cells(shape).get('Relationships'), 'F'), VISIO_RELATIONSHIP_LIST_MEMBERS),
		),
	];
}
const backReferences = (shape: Element, kind: number) =>
	dependsOn(attribute(cells(shape).get('Relationships'), 'F'), kind);
/** Record a list's members: the list names them in order, each member points back with kind 5. */
export function writeListMembers(root: Element, list: Element, memberIds: readonly string[]): void {
	const id = shapeIdOf(list);
	const shapes = new Map(topLevelShapes(root).map((shape) => [shapeIdOf(shape), shape]));
	for (const memberId of listMembersOf(list)) {
		const old = shapes.get(memberId);
		if (old && !memberIds.includes(memberId))
			setRelation(
				old,
				VISIO_RELATIONSHIP_LIST_MEMBER_OF,
				backReferences(old, VISIO_RELATIONSHIP_LIST_MEMBER_OF).filter((other) => other !== id),
			);
	}
	for (const memberId of memberIds) {
		const member = shapes.get(memberId);
		if (!member) continue;
		setRelation(member, VISIO_RELATIONSHIP_LIST_MEMBER_OF, [
			...backReferences(member, VISIO_RELATIONSHIP_LIST_MEMBER_OF).filter((other) => other !== id),
			id,
		]);
		// A container member keeps its (possibly empty) members term, as Visio writes it.
		const terms = relationTerms(attribute(cells(member).get('Relationships'), 'F'));
		if (isContainerSheet(member) && !terms?.some((t) => t.kind === VISIO_RELATIONSHIP_MEMBERS))
			setRelation(member, VISIO_RELATIONSHIP_MEMBERS, []);
	}
	setRelation(list, VISIO_RELATIONSHIP_LIST_MEMBERS, memberIds);
}

/** Every shape that is, directly or through nested containers, a member of `container`. */
export function containerClosure(root: Element, container: Element): Set<string> {
	const shapes = new Map(topLevelShapes(root).map((shape) => [shapeIdOf(shape), shape]));
	const seen = new Set<string>();
	const queue = [container];
	while (queue.length && seen.size < 5000)
		for (const memberId of containerMembersOf(queue.pop()!)) {
			const member = shapes.get(memberId);
			if (!seen.has(memberId)) {
				seen.add(memberId);
				if (member && isContainerSheet(member)) queue.push(member);
			}
		}
	return seen;
}

/** Replace a container's members (the Relationships and back references, not its size). */
export function setVisioContainerMembers(
	root: Element,
	edit: VisioSetContainerMembersEdit,
	check: () => void,
): void {
	check();
	const shapes = topLevelShapes(root);
	const container = shapes.find((shape) => shapeIdOf(shape) === edit.shapeId);
	if (!container || !isContainerSheet(container))
		fail('EDIT_TARGET_NOT_FOUND', 'A top-level container shape is required.');
	const ids = new Set(shapes.map(shapeIdOf));
	for (const memberId of edit.memberIds) {
		if (memberId === edit.shapeId || !ids.has(memberId))
			fail('EDIT_TARGET_NOT_FOUND', 'Container members must be other top-level shapes.');
		const member = shapes.find((shape) => shapeIdOf(shape) === memberId)!;
		if (isContainerSheet(member) && containerClosure(root, member).has(edit.shapeId))
			fail('UNSUPPORTED_GEOMETRY_EDIT', 'A container cannot become a member of its own member.');
	}
	writeContainerMembers(root, container, edit.memberIds);
}

/** Drop references to removed shapes from every container and member (Visio does the same). */
export function pruneContainerReferences(root: Element, removed: ReadonlySet<string>): void {
	for (const shape of topLevelShapes(root)) {
		if (removed.has(shapeIdOf(shape))) continue;
		if (isContainerSheet(shape)) {
			const members = containerMembersOf(shape);
			if (members.some((id) => removed.has(id)))
				writeContainerMembers(
					root,
					shape,
					members.filter((id) => !removed.has(id)),
				);
		}
		const formula = attribute(cells(shape).get('Relationships'), 'F');
		const mentioned = relationTerms(formula)?.some((term) =>
			term.ids.some((id) => removed.has(id)),
		);
		if (!mentioned) continue;
		const next = withoutSheets(formula, removed);
		if (next !== null) writeRelationships(shape, next);
	}
}
