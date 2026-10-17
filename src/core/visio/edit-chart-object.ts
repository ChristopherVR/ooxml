import { NS, parseXml } from '../xml/index';
import { relationshipsPartFor } from '../opc/relationships';
import { RELATIONSHIP_TYPES } from '../opc/relationship-types';
import { VisioPackage } from './package';
import { decodePath, fail, type VisioPackageLimits } from './package-common';
import { related, visioXml } from './parts';
import { attribute, children } from './sheet';
import { openEditablePackage, writeEditedPackage } from './edit-package';
import { serializeEditedXml } from './edit-text';
import { assertShapeIdUnreferenced, cloneRoot, foreignFrameShape } from './edit-picture';
import type {
	VisioChartObjectEdit,
	VisioChartObjectInsertEdit,
	VisioChartObjectReplaceEdit,
} from './edit-chart-object-commands';
import type { EditVsdxResult } from './edit';

const WORKBOOK_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const EMF_TYPE = 'image/x-emf';
/** ForeignData ObjectType of an embedded OLE object Visio draws from a cached preview. */
const OBJECT_TYPE = '33280';

/** Everything one transaction reads and writes; parts are staged here and committed once. */
interface Staged {
	pkg: VisioPackage;
	parts: Map<string, Uint8Array>;
	dirty: Map<string, Element>;
	added: Map<string, Uint8Array>;
	types: Element;
	taken: Set<string>;
}

const lower = (name: string) => decodePath(name).toLowerCase();

/**
 * The first `<stem><n><extension>` that no part already uses, counting up from `start`; a zero
 * suffix is written as none (`Microsoft_Excel_Worksheet.xlsx`, then `...1.xlsx`).
 */
function freeName(
	taken: ReadonlySet<string>,
	directory: string,
	stem: string,
	ext: string,
	start = 0,
): string {
	for (let n = start; ; n++) {
		const name = `${directory}${stem}${n || ''}`;
		if (![...taken].some((part) => part.startsWith(`${name.toLowerCase()}.`)))
			return `${name}.${ext}`;
	}
}

/** Declare `extension` as `contentType`: a Default, or an Override when it already means another type. */
function declareType(types: Element, extension: string, contentType: string, partName: string) {
	const defaults = Array.from(types.getElementsByTagNameNS(types.namespaceURI, 'Default'));
	const declared = defaults.find(
		(node) => node.getAttribute('Extension')?.toLowerCase() === extension,
	);
	if (declared?.getAttribute('ContentType') === contentType) return;
	const entry = types.ownerDocument!.createElementNS(
		types.namespaceURI,
		declared ? 'Override' : 'Default',
	);
	if (declared) entry.setAttribute('PartName', `/${partName}`);
	else entry.setAttribute('Extension', extension);
	entry.setAttribute('ContentType', contentType);
	// Defaults precede overrides in native packages.
	types.insertBefore(entry, declared ? null : (types.firstElementChild ?? null));
}

async function stage(pkg: VisioPackage, parts: Map<string, Uint8Array>): Promise<Staged> {
	const types = cloneRoot(await pkg.readXml('[Content_Types].xml', 'Types'));
	const taken = new Set([...parts.keys()].map(lower));
	for (const node of Array.from(types.getElementsByTagNameNS(types.namespaceURI, 'Override'))) {
		const name = node.getAttribute('PartName');
		if (name?.startsWith('/')) taken.add(lower(name.slice(1)));
	}
	return { pkg, parts, dirty: new Map(), added: new Map(), types, taken };
}

/** Stage the EMF preview of `workbookPart`: its own `image` relationship, as Visio writes it. */
function stagePreview(
	staged: Staged,
	workbookPart: string,
	preview: Uint8Array,
	existing: Element | undefined,
): void {
	const relsPart = relationshipsPartFor(workbookPart);
	const rels = existing ?? parseXml(`<Relationships xmlns="${NS.rels}"/>`).documentElement;
	const media = freeName(staged.taken, 'visio/media/', 'image', 'emf', 1);
	staged.taken.add(media.toLowerCase());
	const ids = new Set(
		Array.from(rels.getElementsByTagNameNS(rels.namespaceURI, 'Relationship')).map(
			(node) => node.getAttribute('Id') ?? '',
		),
	);
	let index = 1;
	while (ids.has(`rId${index}`)) index++;
	const relationship = rels.ownerDocument!.createElementNS(rels.namespaceURI, 'Relationship');
	relationship.setAttribute('Id', `rId${index}`);
	relationship.setAttribute('Type', RELATIONSHIP_TYPES.image);
	relationship.setAttribute('Target', relativeTarget(workbookPart, media));
	rels.appendChild(relationship);
	declareType(staged.types, 'emf', EMF_TYPE, media);
	staged.dirty.set(relsPart, rels);
	staged.added.set(media, preview);
}

/** A relationship Target from the folder of `source` to `part`, as native packages write it. */
function relativeTarget(source: string, part: string): string {
	const from = source.split('/').slice(0, -1);
	const to = part.split('/');
	while (from.length && to.length > 1 && from[0] === to[0]) {
		from.shift();
		to.shift();
	}
	return [...from.map(() => '..'), ...to].join('/');
}

async function insert(
	pkg: VisioPackage,
	parts: Map<string, Uint8Array>,
	pages: ReadonlyMap<string, string>,
	edit: VisioChartObjectInsertEdit,
	check: () => void,
): Promise<{ staged: Staged; pagePart: string }> {
	const pagePart = pages.get(edit.pageId);
	if (!pagePart) fail('EDIT_TARGET_NOT_FOUND', 'Page does not exist.');
	await assertShapeIdUnreferenced(pkg, pages, edit.shapeId, check);
	const staged = await stage(pkg, parts);
	const root = cloneRoot(await visioXml(pkg, pagePart!, 'PageContents'));
	const document = await visioXml(pkg, (await related(pkg, '', 'document'))!, 'VisioDocument');
	const pageRels = relationshipsPartFor(pagePart!);
	const rels = parts.has(pageRels)
		? cloneRoot(await pkg.readXml(pageRels, 'Relationships'))
		: parseXml(`<Relationships xmlns="${NS.rels}"/>`).documentElement;
	const workbookPart = freeName(
		staged.taken,
		'visio/embeddings/',
		'Microsoft_Excel_Worksheet',
		'xlsx',
	);
	staged.taken.add(workbookPart.toLowerCase());
	const ids = new Set(
		Array.from(rels.getElementsByTagNameNS(rels.namespaceURI, 'Relationship')).map(
			(node) => node.getAttribute('Id') ?? '',
		),
	);
	let index = 1;
	while (ids.has(`rId${index}`)) index++;
	const relId = `rId${index}`;
	foreignFrameShape(root, document, edit, relId, {
		ForeignType: 'Object',
		ObjectType: OBJECT_TYPE,
		ShowAsIcon: '0',
		ObjectWidth: String(edit.width),
		ObjectHeight: String(edit.height),
	});
	const relationship = rels.ownerDocument!.createElementNS(rels.namespaceURI, 'Relationship');
	relationship.setAttribute('Id', relId);
	relationship.setAttribute('Type', RELATIONSHIP_TYPES.package);
	relationship.setAttribute('Target', relativeTarget(pagePart!, workbookPart));
	rels.appendChild(relationship);
	declareType(staged.types, 'xlsx', WORKBOOK_TYPE, workbookPart);
	staged.dirty.set(pagePart!, root);
	staged.dirty.set(pageRels, rels);
	staged.added.set(workbookPart, edit.workbook);
	if (edit.preview) stagePreview(staged, workbookPart, edit.preview, undefined);
	return { staged, pagePart: pagePart! };
}

/** The Foreign Object shape `shapeId` on the page, its embedded workbook part and preview part. */
async function locate(
	pkg: VisioPackage,
	pages: ReadonlyMap<string, string>,
	edit: VisioChartObjectReplaceEdit,
): Promise<{ pagePart: string; workbookPart: string; previewPart: string | undefined }> {
	const pagePart = pages.get(edit.pageId);
	if (!pagePart) fail('EDIT_TARGET_NOT_FOUND', 'Page does not exist.');
	const root = await visioXml(pkg, pagePart!, 'PageContents');
	const shape = Array.from(root.getElementsByTagNameNS(root.namespaceURI, 'Shape')).find(
		(node) => node.getAttribute('ID') === edit.shapeId,
	);
	const data = shape && children(shape, 'ForeignData')[0];
	const id = data && children(data, 'Rel')[0]?.getAttributeNS(NS.r, 'id');
	if (
		!shape ||
		attribute(shape, 'Type') !== 'Foreign' ||
		attribute(data, 'ForeignType') !== 'Object' ||
		!id
	)
		fail('EDIT_TARGET_NOT_FOUND', 'The shape is not an embedded chart object.');
	const rel = (await pkg.relationships(pagePart!)).get(id!);
	if (!rel || rel.mode !== 'Internal' || rel.type !== RELATIONSHIP_TYPES.package)
		fail('EDIT_TARGET_NOT_FOUND', 'The shape does not embed a package.');
	const workbookPart = rel!.target;
	const image = [...(await pkg.relationships(workbookPart)).values()].find(
		(candidate) => candidate.type === RELATIONSHIP_TYPES.image && candidate.mode === 'Internal',
	);
	return { pagePart: pagePart!, workbookPart, previewPart: image?.target };
}

async function replace(
	pkg: VisioPackage,
	parts: Map<string, Uint8Array>,
	pages: ReadonlyMap<string, string>,
	edit: VisioChartObjectReplaceEdit,
): Promise<{ staged: Staged; pagePart: string }> {
	const { pagePart, workbookPart, previewPart } = await locate(pkg, pages, edit);
	const staged = await stage(pkg, parts);
	staged.added.set(workbookPart, edit.workbook);
	if (edit.preview) {
		if (previewPart) staged.added.set(previewPart, edit.preview);
		else {
			const relsPart = relationshipsPartFor(workbookPart);
			const existing = parts.has(relsPart)
				? cloneRoot(await pkg.readXml(relsPart, 'Relationships'))
				: undefined;
			stagePreview(staged, workbookPart, edit.preview, existing);
		}
	}
	return { staged, pagePart };
}

/** Insert or replace an embedded chart object (see {@link VisioChartObjectInsertEdit}). */
export async function editVsdxChartObject(
	pkg: VisioPackage,
	parts: Map<string, Uint8Array>,
	pages: ReadonlyMap<string, string>,
	edit: VisioChartObjectEdit,
	limits: VisioPackageLimits,
	maxOutput: number,
	deadline: number,
	check: () => void,
): Promise<EditVsdxResult> {
	const { staged, pagePart } =
		edit.type === 'insert-chart-object'
			? await insert(pkg, parts, pages, edit, check)
			: await replace(pkg, parts, pages, edit);
	staged.dirty.set('[Content_Types].xml', staged.types);
	if (
		new Set([...parts.keys(), ...staged.dirty.keys(), ...staged.added.keys()]).size >
		limits.maxEntries
	)
		fail('LIMIT_ENTRIES', 'Chart insertion exceeds the package entry limit.');
	let total = [...parts.values()].reduce((sum, bytes) => sum + bytes.length, 0);
	for (const [name, bytes] of staged.added) total += bytes.length - (parts.get(name)?.length ?? 0);
	let nodes = 0;
	for (const [name, xml] of staged.dirty) {
		const serialized = serializeEditedXml(xml, limits, check);
		nodes += serialized.nodes;
		if (nodes > limits.maxTotalXmlNodes)
			fail('LIMIT_XML_TOTAL', 'Edited XML node total exceeds limit.');
		total += serialized.bytes.length - (parts.get(name)?.length ?? 0);
		parts.set(name, serialized.bytes);
	}
	if (total > limits.maxTotalBytes) fail('LIMIT_TOTAL', 'Edited package total exceeds limit.');
	for (const [name, bytes] of staged.added) parts.set(name, new Uint8Array(bytes));
	const bytes = await writeEditedPackage(parts, maxOutput, deadline, check);
	const verified = await openEditablePackage(
		bytes,
		{ ...limits, maxInputBytes: maxOutput, maxRuntimeMs: Math.max(1, deadline - Date.now()) },
		check,
	);
	await visioXml(verified.pkg, pagePart, 'PageContents');
	check();
	return {
		bytes,
		changedParts: [...staged.dirty.keys(), ...staged.added.keys()],
		diagnostics: [
			{
				code: 'edit-chart-object-embedded',
				message:
					'The chart was embedded as an Excel workbook object with a metafile preview, as Visio stores Insert > Chart.',
			},
		],
	};
}
