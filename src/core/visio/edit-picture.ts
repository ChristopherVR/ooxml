import { NS, parseXml } from '../xml/index';
import { relationshipsPartFor, resolvePartPath } from '../opc/relationships';
import { RELATIONSHIP_TYPES } from '../opc/relationship-types';
import { VisioPackage } from './package';
import { decodePath, fail, type VisioPackageLimits } from './package-common';
import { related, visioXml } from './parts';
import { attribute } from './sheet';
import { setCell } from './edit-geometry-cells';
import { applyShapeCreationStyles, createShape } from './edit-shape-create';
import { openEditablePackage, writeEditedPackage } from './edit-package';
import { serializeEditedXml } from './edit-text';
import { inspectVisioRasterImage, VisioImageError, type VisioRasterImageInfo } from './media';
import type { VisioPictureInsertEdit } from './edit-metadata-commands';
import type { EditVsdxResult } from './edit';

const FORMATS: Record<
	VisioRasterImageInfo['mimeType'],
	{ extension: string; compression: string }
> = {
	'image/png': { extension: 'png', compression: 'PNG' },
	'image/jpeg': { extension: 'jpeg', compression: 'JPEG' },
	'image/gif': { extension: 'gif', compression: 'GIF' },
};

function copy(root: Element): Element {
	return (root.ownerDocument!.cloneNode(true) as Document).documentElement;
}

/** The new Foreign shape: frame cells, an invisible frame geometry and the bitmap reference. */
function pictureShape(
	root: Element,
	document: Element,
	edit: VisioPictureInsertEdit,
	relId: string,
	compression: string,
): void {
	const shape = createShape(root, edit.shapeId);
	shape.setAttribute('Type', 'Foreign');
	applyShapeCreationStyles(shape, document);
	const node = (name: string) => root.ownerDocument!.createElementNS(root.namespaceURI, name);
	const entries: [string, number, string?][] = [
		['PinX', edit.x],
		['PinY', edit.y],
		['Width', edit.width],
		['Height', edit.height],
		['LocPinX', edit.width / 2, 'Width*0.5'],
		['LocPinY', edit.height / 2, 'Height*0.5'],
		['Angle', 0],
		['FlipX', 0],
		['FlipY', 0],
		['ResizeMode', 0],
		['ImgOffsetX', 0, 'ImgWidth*0'],
		['ImgOffsetY', 0, 'ImgHeight*0'],
		['ImgWidth', edit.width, 'Width*1'],
		['ImgHeight', edit.height, 'Height*1'],
	];
	for (const [name, value, formula] of entries) setCell(shape, name, value, formula);
	const section = node('Section');
	section.setAttribute('N', 'Geometry');
	section.setAttribute('IX', '0');
	for (const [name, value] of [
		['NoFill', 1],
		['NoLine', 1],
		['NoShow', 0],
		['NoSnap', 0],
	] as const)
		setCell(section, name, value);
	for (const [index, [x, y]] of [
		[0, 0],
		[1, 0],
		[1, 1],
		[0, 1],
		[0, 0],
	].entries()) {
		const row = node('Row');
		row.setAttribute('T', index ? 'RelLineTo' : 'RelMoveTo');
		row.setAttribute('IX', String(index + 1));
		setCell(row, 'X', x!);
		setCell(row, 'Y', y!);
		section.appendChild(row);
	}
	shape.appendChild(section);
	const foreign = node('ForeignData');
	foreign.setAttribute('ForeignType', 'Bitmap');
	foreign.setAttribute('CompressionType', compression);
	const rel = node('Rel');
	rel.setAttributeNS(NS.r, 'r:id', relId);
	foreign.appendChild(rel);
	shape.appendChild(foreign);
}

/** Embed one validated raster as visio/media part, page relationship and Foreign shape. */
export async function editVsdxPicture(
	pkg: VisioPackage,
	parts: Map<string, Uint8Array>,
	pages: ReadonlyMap<string, string>,
	edit: VisioPictureInsertEdit,
	limits: VisioPackageLimits,
	maxOutput: number,
	deadline: number,
	check: () => void,
): Promise<EditVsdxResult> {
	let info: VisioRasterImageInfo;
	try {
		info = inspectVisioRasterImage(edit.image);
	} catch (error) {
		if (error instanceof VisioImageError) fail('INVALID_PICTURE', error.message);
		throw error;
	}
	const format = FORMATS[info.mimeType];
	const path = pages.get(edit.pageId);
	if (!path) fail('EDIT_TARGET_NOT_FOUND', 'Page does not exist.');
	// A dangling Sheet.N! reference must not start resolving to the new picture.
	const reference = new RegExp(`\\bSheet\\.${edit.shapeId}!`, 'i');
	for (const pagePath of pages.values()) {
		const pageRoot = await visioXml(pkg, pagePath, 'PageContents');
		for (const cell of Array.from(pageRoot.getElementsByTagNameNS(pageRoot.namespaceURI, 'Cell'))) {
			check();
			if (reference.test(attribute(cell, 'F') ?? ''))
				fail('INVALID_SHAPE_ID', 'An existing formula already refers to the new shape ID.');
		}
	}
	const root = copy(await visioXml(pkg, path!, 'PageContents'));
	const document = await visioXml(pkg, (await related(pkg, '', 'document'))!, 'VisioDocument');
	const relsPart = relationshipsPartFor(path!);
	const rels = parts.has(relsPart)
		? copy(await pkg.readXml(relsPart, 'Relationships'))
		: parseXml(`<Relationships xmlns="${NS.rels}"/>`).documentElement;
	const types = copy(await pkg.readXml('[Content_Types].xml', 'Types'));
	const taken = new Set([...parts.keys()].map((name) => decodePath(name).toLowerCase()));
	const ids = new Set<string>();
	for (const node of Array.from(types.getElementsByTagNameNS(types.namespaceURI, 'Override'))) {
		const name = node.getAttribute('PartName');
		if (name?.startsWith('/')) taken.add(decodePath(name.slice(1)).toLowerCase());
	}
	for (const node of Array.from(rels.getElementsByTagNameNS(rels.namespaceURI, 'Relationship')))
		ids.add(node.getAttribute('Id') ?? '');
	let index = 1;
	let target = '';
	let media = '';
	do {
		check();
		target = `../media/image${index++}.${format.extension}`;
		media = resolvePartPath(path!, target);
	} while ([...taken].some((name) => name.startsWith(media.toLowerCase().replace(/[^.]+$/, ''))));
	index = 1;
	while (ids.has(`rId${index}`)) index++;
	const relId = `rId${index}`;
	pictureShape(root, document, edit, relId, format.compression);
	const relationship = rels.ownerDocument!.createElementNS(rels.namespaceURI, 'Relationship');
	relationship.setAttribute('Id', relId);
	relationship.setAttribute('Type', RELATIONSHIP_TYPES.image);
	relationship.setAttribute('Target', target);
	rels.appendChild(relationship);
	const defaults = Array.from(types.getElementsByTagNameNS(types.namespaceURI, 'Default'));
	const declared = defaults.find(
		(node) => node.getAttribute('Extension')?.toLowerCase() === format.extension,
	);
	if (!declared || declared.getAttribute('ContentType') !== info.mimeType) {
		const entry = types.ownerDocument!.createElementNS(
			types.namespaceURI,
			declared ? 'Override' : 'Default',
		);
		if (declared) entry.setAttribute('PartName', `/${media}`);
		else entry.setAttribute('Extension', format.extension);
		entry.setAttribute('ContentType', info.mimeType);
		// Defaults precede overrides in native packages.
		types.insertBefore(entry, declared ? null : (types.firstElementChild ?? null));
	}
	const dirty = new Map<string, Element>([
		[path!, root],
		[relsPart, rels],
		['[Content_Types].xml', types],
	]);
	if (new Set([...parts.keys(), ...dirty.keys(), media]).size > limits.maxEntries)
		fail('LIMIT_ENTRIES', 'Picture insertion exceeds the package entry limit.');
	let total = [...parts.values()].reduce((sum, bytes) => sum + bytes.length, 0) + edit.image.length;
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
	parts.set(media, new Uint8Array(edit.image));
	const bytes = await writeEditedPackage(parts, maxOutput, deadline, check);
	const verified = await openEditablePackage(
		bytes,
		{ ...limits, maxInputBytes: maxOutput, maxRuntimeMs: Math.max(1, deadline - Date.now()) },
		check,
	);
	await visioXml(verified.pkg, path!, 'PageContents');
	check();
	return {
		bytes,
		changedParts: [...dirty.keys(), media],
		diagnostics: [
			{
				code: 'edit-picture-experimental',
				message:
					'The picture was embedded as a Foreign bitmap shape with a media part. Native Visio reopen fidelity is limited to tested cases.',
			},
		],
	};
}
