import { relationshipsPartFor, resolvePartPath } from '../opc/relationships';
import { NS, parseXml } from '../xml/index';
import { calloutInstanceMarkup, calloutMasterName } from './data-graphic-instance';
import {
	calloutMasterSpec,
	dataGraphicMasterSpec,
	importMarkup,
	userRow,
	type NativeDataGraphicKind,
	type NativeMasterItem,
} from './data-graphic-masters';
import {
	VISIO_DATA_GRAPHIC_DEFAULT_POSITION,
	VISIO_DATA_GRAPHIC_POSITIONS,
	visioDataGraphicPoint,
	type VisioDataGraphicPosition,
} from './data-graphic-place';
import type { EditVsdxResult } from './edit';
import { addMasterToDrawing, shapesOf } from './edit-master-add';
import { cells, numeric, setCell } from './edit-geometry-cells';
import { openEditablePackage, writeEditedPackage } from './edit-package';
import { serializeEditedXml } from './edit-text';
import { fail, type VisioPackageLimits } from './package-common';
import { indexedPart, related, visioXml } from './parts';
import { attribute, children } from './sheet';

/** One callout of a native data graphic. */
export interface VisioDataGraphicItemSpec {
	kind: NativeDataGraphicKind;
	/** The owner's Shape Data row (its name, not its label) the callout reads. */
	field: string;
	position?: VisioDataGraphicPosition;
	/** The range of a bar or an icon set; the owner's values decide it when absent. */
	min?: number;
	max?: number;
}

/**
 * Make a shape's data graphic exactly these callouts: Visio's structure for it (the shape
 * becomes a group, each callout an instance of a callout master that reads the shape's Shape
 * Data by formula, the graphic a hidden data graphic master named by the shape's
 * `User.visDGDisplayFormat`). An empty list removes the graphic. The shape must be a local
 * shape; an instance of a stencil master is refused.
 */
export interface VisioSetDataGraphicEdit {
	type: 'set-data-graphic';
	pageId: string;
	shapeId: string;
	items: VisioDataGraphicItemSpec[];
}

export const isVisioDataGraphicEdit = (edit: { type: string }): edit is VisioSetDataGraphicEdit =>
	edit.type === 'set-data-graphic';

const ID = /^[1-9]\d{0,9}$/;
const MAX_ITEMS = 24;

export function snapshotDataGraphicEdit(edit: VisioSetDataGraphicEdit): VisioSetDataGraphicEdit {
	if (typeof edit.pageId !== 'string' || !edit.pageId || edit.pageId.length > 256)
		fail('INVALID_EDIT', 'Invalid edit target.');
	if (typeof edit.shapeId !== 'string' || !ID.test(edit.shapeId) || Number(edit.shapeId) > 2 ** 31)
		fail('INVALID_EDIT', 'Invalid edit shape target.');
	if (!Array.isArray(edit.items) || edit.items.length > MAX_ITEMS)
		fail('INVALID_EDIT', 'A data graphic has at most 24 items.');
	return {
		type: 'set-data-graphic',
		pageId: edit.pageId,
		shapeId: edit.shapeId,
		items: edit.items.map((item) => {
			if (!['text', 'bar', 'icon'].includes(item.kind))
				fail('INVALID_EDIT', 'Unknown data graphic item kind.');
			if (typeof item.field !== 'string' || !/^[\p{L}_][\p{L}\p{N}_]{0,127}$/u.test(item.field))
				fail('INVALID_EDIT', 'A data graphic item reads a Shape Data row by name.');
			if (item.position !== undefined && !VISIO_DATA_GRAPHIC_POSITIONS.includes(item.position))
				fail('INVALID_EDIT', 'Unknown data graphic position.');
			for (const value of [item.min, item.max])
				if (value !== undefined && (typeof value !== 'number' || !Number.isFinite(value)))
					fail('INVALID_EDIT', 'A data graphic range must be finite numbers.');
			return {
				kind: item.kind,
				field: item.field,
				...(item.position ? { position: item.position } : {}),
				...(item.min === undefined ? {} : { min: item.min }),
				...(item.max === undefined ? {} : { max: item.max }),
			};
		}),
	};
}

const copy = (root: Element): Element =>
	(root.ownerDocument!.cloneNode(true) as Document).documentElement;
/** The rows the graphic adds to its owner's User section. */
const OWNER_ROWS = ['visDGDisplayFormat', 'msvLayoutIncludeSubshapes'];
const MASTER_REL = 'http://schemas.microsoft.com/visio/2010/relationships/master';

const fnv = (text: string): string => {
	let hash = 0x811c9dc5;
	for (let index = 0; index < text.length; index++)
		hash = Math.imul(hash ^ text.charCodeAt(index), 0x01000193) >>> 0;
	return hash.toString(16).toUpperCase().padStart(8, '0');
};

function findShape(root: Element, id: string): Element | undefined {
	const pending = shapesOf(root);
	while (pending.length) {
		const shape = pending.shift()!;
		if (attribute(shape, 'ID') === id) return shape;
		pending.push(...shapesOf(shape));
	}
	return undefined;
}

const isCallout = (shape: Element): boolean =>
	children(shape, 'Section').some(
		(section) =>
			attribute(section, 'N') === 'User' &&
			children(section, 'Row').some((row) => attribute(row, 'N') === 'visDGField0'),
	);

/** Take the graphic's callouts and its owner rows off a shape, as far as it was made by this. */
function clear(owner: Element): string[] {
	const removed: string[] = [];
	for (const container of children(owner, 'Shapes'))
		for (const member of children(container, 'Shape'))
			if (isCallout(member)) {
				removed.push(attribute(member, 'ID') ?? '');
				container.removeChild(member);
			}
	for (const container of children(owner, 'Shapes'))
		if (!children(container, 'Shape').length) owner.removeChild(container);
	for (const section of children(owner, 'Section').filter(
		(item) => attribute(item, 'N') === 'User',
	)) {
		for (const row of children(section, 'Row'))
			if (OWNER_ROWS.includes(attribute(row, 'N') ?? '')) section.removeChild(row);
		if (!children(section, 'Row').length) owner.removeChild(section);
	}
	if (attribute(owner, 'Type') === 'Group' && !children(owner, 'Shapes').length) {
		owner.setAttribute('Type', 'Shape');
		const mode = cells(owner).get('DisplayMode');
		if (mode) owner.removeChild(mode);
	}
	return removed;
}

interface Row {
	label: string;
	text: string;
	number?: number;
}

function propertyOf(owner: Element, name: string): Row {
	const section = children(owner, 'Section').find((item) => attribute(item, 'N') === 'Property');
	const row = section && children(section, 'Row').find((item) => attribute(item, 'N') === name);
	if (!row) fail('EDIT_TARGET_NOT_FOUND', `The shape has no Shape Data row named ${name}.`);
	const rowCells = cells(row!);
	const value = rowCells.get('Value');
	const text = attribute(value!, 'V') ?? '';
	const label = attribute(rowCells.get('Label') ?? row!, 'V') ?? '';
	const asNumber = text !== '' && attribute(value!, 'U') !== 'STR' ? Number(text) : NaN;
	return {
		label: label || name,
		text,
		...(Number.isFinite(asNumber) ? { number: asNumber } : {}),
	};
}

function insertSection(shape: Element, section: Element): void {
	const anchor = Array.from(shape.childNodes).find(
		(child) =>
			child.nodeType === 1 &&
			['Text', 'Data1', 'Data2', 'Data3', 'ForeignData', 'Shapes'].includes(
				(child as Element).localName,
			),
	);
	shape.insertBefore(section, anchor ?? null);
}

/** The package-level plan of one edit: which masters the page needs, then the page change. */
async function plan(
	bytes: Uint8Array,
	edit: VisioSetDataGraphicEdit,
	limits: VisioPackageLimits,
	check: () => void,
) {
	const { pkg, pages } = await openEditablePackage(bytes, limits, check);
	const path = pages.get(edit.pageId);
	if (!path) fail('EDIT_TARGET_NOT_FOUND', 'Page does not exist.');
	const root = await visioXml(pkg, path!, 'PageContents');
	const owner = findShape(root, edit.shapeId);
	if (!owner) fail('EDIT_TARGET_NOT_FOUND', 'Shape does not exist.');
	if (owner!.hasAttribute('Master') || owner!.hasAttribute('MasterShape'))
		fail(
			'UNSUPPORTED_DATA_GRAPHIC',
			'A data graphic cannot be applied to an instance of a master yet.',
		);
	if (!['Shape', 'Group'].includes(attribute(owner!, 'Type') ?? 'Shape'))
		fail('UNSUPPORTED_DATA_GRAPHIC', 'A data graphic needs a shape or a group.');
	for (const item of edit.items) propertyOf(owner!, item.field);
}

export async function editVsdxDataGraphics(
	source: Uint8Array,
	edits: readonly VisioSetDataGraphicEdit[],
	limits: VisioPackageLimits,
	maxOutput: number,
	deadline: number,
	check: () => void,
): Promise<EditVsdxResult> {
	let bytes = source;
	const changed = new Set<string>();
	for (const edit of edits) {
		check();
		await plan(bytes, edit, limits, check);
		const masters = new Map<NativeDataGraphicKind, string>();
		let graphicId = '';
		let graphicUid = '';
		if (edit.items.length) {
			for (const kind of new Set(edit.items.map((item) => item.kind))) {
				const added = await addMasterToDrawing(
					bytes,
					calloutMasterSpec(kind),
					limits,
					maxOutput,
					deadline,
					check,
				);
				bytes = added.bytes;
				masters.set(kind, added.masterId);
				for (const part of added.changedParts) changed.add(part);
			}
			const listed: NativeMasterItem[] = edit.items.map((item) => {
				const point = visioDataGraphicPoint(
					item.position ?? VISIO_DATA_GRAPHIC_DEFAULT_POSITION[item.kind],
				);
				return { kind: item.kind, field: item.field, position: `PNT(${point.x},${point.y})` };
			});
			graphicUid = `{${fnv(listed.map((item) => `${item.kind}:${item.field}:${item.position}`).join('|'))}-0000-0000-8E40-00608CF305B2}`;
			const added = await addMasterToDrawing(
				bytes,
				dataGraphicMasterSpec(graphicUid, 'Data Graphic', listed),
				limits,
				maxOutput,
				deadline,
				check,
			);
			bytes = added.bytes;
			graphicId = added.masterId;
			for (const part of added.changedParts) changed.add(part);
		}
		const done = await writePage(
			bytes,
			edit,
			masters,
			graphicUid,
			graphicId,
			limits,
			maxOutput,
			deadline,
			check,
		);
		bytes = done.bytes;
		for (const part of done.changedParts) changed.add(part);
	}
	return { bytes, changedParts: [...changed], diagnostics: [] };
}

async function writePage(
	bytes: Uint8Array,
	edit: VisioSetDataGraphicEdit,
	masters: ReadonlyMap<NativeDataGraphicKind, string>,
	graphicUid: string,
	graphicId: string,
	limits: VisioPackageLimits,
	maxOutput: number,
	deadline: number,
	check: () => void,
): Promise<{ bytes: Uint8Array; changedParts: string[] }> {
	const { pkg, parts, pages } = await openEditablePackage(bytes, limits, check);
	const path = pages.get(edit.pageId)!;
	const root = copy(await visioXml(pkg, path, 'PageContents'));
	const owner = findShape(root, edit.shapeId)!;
	// Callouts keep their IDs when the graphic is rewritten, so a selection survives a refresh.
	const reusable = clear(owner);
	const taken = new Set<string>();
	const names = new Set<string>();
	const pending = shapesOf(root);
	while (pending.length) {
		const shape = pending.pop()!;
		taken.add(attribute(shape, 'ID') ?? '');
		names.add(attribute(shape, 'NameU') ?? '');
		pending.push(...shapesOf(shape));
	}
	const dirty = new Map<string, Element>([[path, root]]);
	if (edit.items.length) {
		const ownerCells = cells(owner);
		const width = numeric(ownerCells.get('Width')),
			height = numeric(ownerCells.get('Height'));
		const values = edit.items.map((item) => propertyOf(owner, item.field).number);
		let next = 0;
		const fresh = () => {
			const old = reusable.shift();
			if (old) {
				taken.add(old);
				return old;
			}
			do next++;
			while (taken.has(String(next)));
			taken.add(String(next));
			return String(next);
		};
		owner.setAttribute('Type', 'Group');
		// A group's pose edits need the cells a Visio-saved shape always has.
		for (const [name, value, formula] of [
			['LocPinX', width / 2, 'Width*0.5'],
			['LocPinY', height / 2, 'Height*0.5'],
			['Angle', 0],
			['FlipX', 0],
			['FlipY', 0],
		] as const)
			if (!ownerCells.has(name)) setCell(owner, name, value, formula);
		if (!ownerCells.has('DisplayMode')) {
			const mode = root.ownerDocument!.createElementNS(root.namespaceURI, 'Cell');
			mode.setAttribute('N', 'DisplayMode');
			mode.setAttribute('V', '1');
			owner.insertBefore(
				mode,
				Array.from(owner.childNodes).find(
					(child) => child.nodeType === 1 && (child as Element).localName !== 'Cell',
				) ?? null,
			);
		}
		const fromUser = children(owner, 'Section').find(
			(section) => attribute(section, 'N') === 'User',
		);
		const rows = [
			userRow('visDGDisplayFormat', 254, `USE(${graphicUid})`),
			userRow('msvLayoutIncludeSubshapes', 1, 'TRUE', 'BOOL'),
		].join('');
		const document = root.ownerDocument!;
		if (fromUser) {
			const holder = importMarkup(
				document,
				root.namespaceURI!,
				`<Section N='User'>${rows}</Section>`,
			);
			for (const row of children(holder, 'Row')) fromUser.appendChild(row);
		} else
			insertSection(
				owner,
				importMarkup(document, root.namespaceURI!, `<Section N='User'>${rows}</Section>`),
			);
		let container = children(owner, 'Shapes')[0];
		if (!container) {
			container = document.createElementNS(root.namespaceURI, 'Shapes');
			owner.appendChild(container);
		}
		const used = new Map<string, number>();
		for (const [index, item] of edit.items.entries()) {
			const position = item.position ?? VISIO_DATA_GRAPHIC_DEFAULT_POSITION[item.kind];
			const slot = used.get(position) ?? 0;
			used.set(position, slot + 1);
			const row = propertyOf(owner, item.field);
			const known = values.filter((value): value is number => value !== undefined);
			const min = item.min ?? Math.min(0, ...known);
			const max = item.max ?? Math.max(...known, min + 1);
			const id = fresh();
			const base = calloutMasterName(item.kind);
			const name = names.has(base) ? `${base}.${id}` : base;
			names.add(name);
			container.appendChild(
				importMarkup(
					document,
					root.namespaceURI!,
					calloutInstanceMarkup({
						id,
						name,
						masterId: masters.get(item.kind)!,
						ownerId: edit.shapeId,
						kind: item.kind,
						field: item.field,
						label: row.label,
						value: { text: row.text, ...(row.number === undefined ? {} : { number: row.number }) },
						position,
						slot,
						owner: { width, height },
						min,
						max,
					}),
				),
			);
			void index;
		}
		const relsPart = relationshipsPartFor(path);
		const rels = parts.has(relsPart)
			? copy(await pkg.readXml(relsPart, 'Relationships'))
			: parseXml(`<Relationships xmlns="${NS.rels}"/>`).documentElement;
		const existing = Array.from(rels.getElementsByTagNameNS(rels.namespaceURI, 'Relationship'));
		const documentPart = (await related(pkg, '', 'document'))!;
		const mastersPart = (await related(pkg, documentPart, 'masters'))!;
		const mastersRoot = await visioXml(pkg, mastersPart, 'Masters');
		let changedRels = false;
		for (const masterId of new Set(masters.values())) {
			const entry = children(mastersRoot, 'Master').find(
				(item) => attribute(item, 'ID') === masterId,
			)!;
			const masterPart = await indexedPart(pkg, mastersPart, entry, 'master');
			const linked = existing.some(
				(rel) =>
					rel.getAttribute('Type') === MASTER_REL &&
					rel.getAttribute('TargetMode') !== 'External' &&
					resolvePartPath(path, rel.getAttribute('Target') ?? '') === masterPart,
			);
			if (linked) continue;
			const ids = new Set(
				Array.from(rels.getElementsByTagNameNS(rels.namespaceURI, 'Relationship')).map(
					(rel) => rel.getAttribute('Id') ?? '',
				),
			);
			let index = 1;
			while (ids.has(`rId${index}`)) index++;
			const relationship = rels.ownerDocument!.createElementNS(rels.namespaceURI, 'Relationship');
			relationship.setAttribute('Id', `rId${index}`);
			relationship.setAttribute('Type', MASTER_REL);
			const from = path.split('/').slice(0, -1);
			const to = masterPart.split('/');
			let shared = 0;
			while (shared < from.length && shared < to.length - 1 && from[shared] === to[shared])
				shared++;
			relationship.setAttribute(
				'Target',
				[...from.slice(shared).map(() => '..'), ...to.slice(shared)].join('/'),
			);
			rels.appendChild(relationship);
			changedRels = true;
		}
		if (changedRels) dirty.set(relsPart, rels);
		void graphicId;
	}
	let total = [...parts.values()].reduce((sum, part) => sum + part.length, 0);
	let nodes = 0;
	if (new Set([...parts.keys(), ...dirty.keys()]).size > limits.maxEntries)
		fail('LIMIT_ENTRIES', 'The edit exceeds the package entry limit.');
	for (const [name, xml] of dirty) {
		const serialized = serializeEditedXml(xml, limits, check);
		nodes += serialized.nodes;
		if (nodes > limits.maxTotalXmlNodes)
			fail('LIMIT_XML_TOTAL', 'Edited XML node total exceeds limit.');
		total += serialized.bytes.length - (parts.get(name)?.length ?? 0);
		if (total > limits.maxTotalBytes) fail('LIMIT_TOTAL', 'Edited package total exceeds limit.');
		parts.set(name, serialized.bytes);
	}
	const output = await writeEditedPackage(parts, maxOutput, deadline, check);
	const verified = await openEditablePackage(
		output,
		{ ...limits, maxInputBytes: maxOutput, maxRuntimeMs: Math.max(1, deadline - Date.now()) },
		check,
	);
	await visioXml(verified.pkg, path, 'PageContents');
	return { bytes: output, changedParts: [...dirty.keys()] };
}
