import { NS, parseXml } from '../xml/index';
import { relationshipsPartFor, resolvePartPath } from '../opc/relationships';
import { VisioPackage } from './package';
import { fail, type VisioPackageLimits } from './package-common';
import { indexedPart, related, visioXml } from './parts';
import { attribute, children } from './sheet';
import { setCell } from './edit-geometry-cells';
import { createShape } from './edit-shape-create';
import { openEditablePackage, writeEditedPackage } from './edit-package';
import { serializeEditedXml } from './edit-text';
import type { EditVsdxResult } from './edit';

const MASTER_RELATIONSHIP = 'http://schemas.microsoft.com/visio/2010/relationships/master';
/** Sub-shapes of one instance; a larger group master is refused rather than half-instanced. */
const MAX_INSTANCE_SHAPES = 2_000;

/**
 * Drop a master of the drawing's document stencil on a page, as Visio does: a shape that names the
 * master and carries only its pin, so size, geometry, text and formatting stay inherited. Drawing
 * inches, bottom-left origin; `x`/`y` is where the master's pin lands.
 */
export interface VisioMasterInstanceEdit {
	type: 'insert-master-instance';
	pageId: string;
	/** ID of the new shape; sub-shapes of a group master take the next free IDs. */
	shapeId: string;
	masterId: string;
	x: number;
	y: number;
}

const ID = /^[1-9]\d{0,9}$/;

/** Copy and validate the command; page and master admission stay with the transaction. */
export function snapshotMasterInstance(edit: VisioMasterInstanceEdit): VisioMasterInstanceEdit {
	if (typeof edit.pageId !== 'string' || !edit.pageId || edit.pageId.length > 256)
		fail('INVALID_EDIT', 'Invalid edit target.');
	if (typeof edit.shapeId !== 'string' || !ID.test(edit.shapeId) || Number(edit.shapeId) > 2 ** 31)
		fail('INVALID_EDIT', 'Invalid edit shape target.');
	if (typeof edit.masterId !== 'string' || !/^(0|[1-9]\d{0,9})$/.test(edit.masterId))
		fail('INVALID_EDIT', 'Master IDs must be canonical unsigned integers.');
	for (const value of [edit.x, edit.y])
		if (typeof value !== 'number' || !Number.isFinite(value) || Math.abs(value) > 1e6)
			fail('INVALID_EDIT', 'The drop point must be finite drawing inches within limits.');
	return {
		type: 'insert-master-instance',
		pageId: edit.pageId,
		shapeId: edit.shapeId,
		masterId: edit.masterId,
		x: edit.x,
		y: edit.y,
	};
}

function copy(root: Element): Element {
	return (root.ownerDocument!.cloneNode(true) as Document).documentElement;
}

const shapesOf = (parent: Element): Element[] =>
	children(parent, 'Shapes').flatMap((container) => children(container, 'Shape'));

/** Every shape ID on the page, at any depth. */
function pageIds(root: Element): Set<string> {
	const ids = new Set<string>();
	const pending = [root];
	while (pending.length) {
		const parent = pending.pop()!;
		for (const shape of shapesOf(parent)) {
			ids.add(attribute(shape, 'ID') ?? '');
			pending.push(shape);
		}
	}
	return ids;
}

/** Add the instance shape, its inherited sub-shapes and the page's relationship to the master. */
export async function editVsdxMasterInstance(
	pkg: VisioPackage,
	parts: Map<string, Uint8Array>,
	pages: ReadonlyMap<string, string>,
	edit: VisioMasterInstanceEdit,
	limits: VisioPackageLimits,
	maxOutput: number,
	deadline: number,
	check: () => void,
): Promise<EditVsdxResult> {
	const path = pages.get(edit.pageId);
	if (!path) fail('EDIT_TARGET_NOT_FOUND', 'Page does not exist.');
	const documentPart = (await related(pkg, '', 'document'))!;
	const mastersPart = await related(pkg, documentPart, 'masters', false);
	const master = mastersPart
		? children(await visioXml(pkg, mastersPart, 'Masters'), 'Master').filter(
				(node) => attribute(node, 'ID') === edit.masterId,
			)
		: [];
	if (!mastersPart || master.length !== 1)
		fail('EDIT_TARGET_NOT_FOUND', 'The master does not belong to this drawing.');
	const masterPart = await indexedPart(pkg, mastersPart!, master[0]!, 'master');
	const roots = shapesOf(await visioXml(pkg, masterPart, 'MasterContents')).filter(
		(shape) => attribute(shape, 'Del') !== '1',
	);
	if (roots.length !== 1)
		fail(
			'UNSUPPORTED_MASTER_INSTANCE',
			'Only a master with exactly one top-level shape can be dropped.',
		);
	const base = roots[0]!;
	const cellNames = new Set(children(base, 'Cell').map((cell) => attribute(cell, 'N')));
	if (cellNames.has('BeginX') || cellNames.has('EndX'))
		fail(
			'UNSUPPORTED_MASTER_INSTANCE',
			'A 1-D master (a connector or a line) cannot be dropped as a shape.',
		);
	if (base.hasAttribute('Master') || base.hasAttribute('MasterShape'))
		fail('UNSUPPORTED_MASTER_INSTANCE', 'The master inherits another master.');
	const type = attribute(base, 'Type') ?? 'Shape';
	if (type !== 'Shape' && type !== 'Group')
		fail('UNSUPPORTED_MASTER_INSTANCE', `A ${type} master cannot be dropped.`);

	// A dangling Sheet.N! reference must not start resolving to a new shape.
	const formulas: string[] = [];
	for (const pagePath of pages.values()) {
		const pageRoot = await visioXml(pkg, pagePath, 'PageContents');
		for (const cell of Array.from(pageRoot.getElementsByTagNameNS(pageRoot.namespaceURI, 'Cell'))) {
			check();
			const formula = attribute(cell, 'F');
			if (formula?.includes('!')) formulas.push(formula);
		}
	}
	const root = copy(await visioXml(pkg, path!, 'PageContents'));
	const taken = pageIds(root);
	const claim = (id: string): string => {
		if (taken.has(id)) fail('INVALID_SHAPE_ID', 'Shape ID already exists.');
		const reference = new RegExp(`\\bSheet\\.${id}!`, 'i');
		if (formulas.some((formula) => reference.test(formula)))
			fail('INVALID_SHAPE_ID', 'An existing formula already refers to the new shape ID.');
		taken.add(id);
		return id;
	};
	let next = Number(edit.shapeId);
	const fresh = (): string => {
		do next++;
		while (taken.has(String(next)));
		if (next > 2 ** 31) fail('INVALID_SHAPE_ID', 'The page has no free shape IDs.');
		return claim(String(next));
	};
	const shape = createShape(root, claim(edit.shapeId));
	shape.setAttribute('Type', type);
	shape.setAttribute('Master', edit.masterId);
	setCell(shape, 'PinX', edit.x);
	setCell(shape, 'PinY', edit.y);
	const node = (name: string) => root.ownerDocument!.createElementNS(root.namespaceURI, name);
	let count = 0;
	const inherit = (from: Element, to: Element): void => {
		const members = shapesOf(from).filter((member) => attribute(member, 'Del') !== '1');
		if (!members.length) return;
		const container = node('Shapes');
		to.appendChild(container);
		for (const member of members) {
			check();
			if (++count > MAX_INSTANCE_SHAPES)
				fail('UNSUPPORTED_MASTER_INSTANCE', 'The master has too many sub-shapes to drop.');
			const id = attribute(member, 'ID');
			if (!id || member.hasAttribute('Master'))
				fail('UNSUPPORTED_MASTER_INSTANCE', 'The master nests another master.');
			const child = node('Shape');
			child.setAttribute('ID', fresh());
			child.setAttribute('Type', attribute(member, 'Type') ?? 'Shape');
			child.setAttribute('MasterShape', id!);
			container.appendChild(child);
			inherit(member, child);
		}
	};
	inherit(base, shape);

	const relsPart = relationshipsPartFor(path!);
	const rels = parts.has(relsPart)
		? copy(await pkg.readXml(relsPart, 'Relationships'))
		: parseXml(`<Relationships xmlns="${NS.rels}"/>`).documentElement;
	const existing = Array.from(rels.getElementsByTagNameNS(rels.namespaceURI, 'Relationship'));
	const linked = existing.some(
		(rel) =>
			rel.getAttribute('Type') === MASTER_RELATIONSHIP &&
			rel.getAttribute('TargetMode') !== 'External' &&
			resolvePartPath(path!, rel.getAttribute('Target') ?? '') === masterPart,
	);
	const dirty = new Map<string, Element>([[path!, root]]);
	if (!linked) {
		const ids = new Set(existing.map((rel) => rel.getAttribute('Id') ?? ''));
		let index = 1;
		while (ids.has(`rId${index}`)) index++;
		const relationship = rels.ownerDocument!.createElementNS(rels.namespaceURI, 'Relationship');
		relationship.setAttribute('Id', `rId${index}`);
		relationship.setAttribute('Type', MASTER_RELATIONSHIP);
		// Page and master parts are siblings under visio/ in every package Visio writes.
		const from = path!.split('/').slice(0, -1);
		const to = masterPart.split('/');
		let shared = 0;
		while (shared < from.length && shared < to.length - 1 && from[shared] === to[shared]) shared++;
		relationship.setAttribute(
			'Target',
			[...from.slice(shared).map(() => '..'), ...to.slice(shared)].join('/'),
		);
		rels.appendChild(relationship);
		dirty.set(relsPart, rels);
	}
	if (new Set([...parts.keys(), ...dirty.keys()]).size > limits.maxEntries)
		fail('LIMIT_ENTRIES', 'The edit exceeds the package entry limit.');
	let total = [...parts.values()].reduce((sum, bytes) => sum + bytes.length, 0);
	let nodes = 0;
	for (const [name, xml] of dirty) {
		const serialized = serializeEditedXml(xml, limits, check);
		nodes += serialized.nodes;
		if (nodes > limits.maxTotalXmlNodes)
			fail('LIMIT_XML_TOTAL', 'Edited XML node total exceeds limit.');
		total += serialized.bytes.length - (parts.get(name)?.length ?? 0);
		if (total > limits.maxTotalBytes) fail('LIMIT_TOTAL', 'Edited package total exceeds limit.');
		parts.set(name, serialized.bytes);
	}
	const bytes = await writeEditedPackage(parts, maxOutput, deadline, check);
	const verified = await openEditablePackage(
		bytes,
		{ ...limits, maxInputBytes: maxOutput, maxRuntimeMs: Math.max(1, deadline - Date.now()) },
		check,
	);
	await visioXml(verified.pkg, path!, 'PageContents');
	check();
	return { bytes, changedParts: [...dirty.keys()], diagnostics: [] };
}
