import { attribute, child, children } from './sheet.js';
import { related, visioXml } from './parts.js';
import { VisioPackage } from './package.js';
import { decodePath, fail, type VisioPackageLimits } from './package-common.js';
import { openEditablePackage, writeEditedPackage } from './edit-package.js';
import { serializeEditedXml } from './edit-text.js';
import { parseXml } from '../xml/index.js';
import { updatePageAppProperties } from './edit-page-properties.js';
import type { VisioPageInsert } from './edit-commands.js';
import type { EditVsdxResult } from './edit.js';

const officeRel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const pageType = 'http://schemas.microsoft.com/visio/2010/relationships/page';
const contentType = 'application/vnd.ms-visio.page+xml';
function copy(root: Element): Element {
	return (root.ownerDocument!.cloneNode(true) as Document).documentElement;
}

/** Source-backed page metadata/OPC transaction. Existing page payloads stay untouched. */
export async function insertVsdxPages(
	pkg: VisioPackage,
	parts: Map<string, Uint8Array>,
	commands: readonly VisioPageInsert[],
	limits: VisioPackageLimits,
	maxOutput: number,
	deadline: number,
	check: () => void,
): Promise<EditVsdxResult> {
	const documentPart = (await related(pkg, '', 'document'))!;
	const pagesPart = (await related(pkg, documentPart, 'pages'))!;
	const slash = pagesPart.lastIndexOf('/');
	const directory = pagesPart.slice(0, slash + 1);
	const relsPart = `${directory}_rels/${pagesPart.slice(slash + 1)}.rels`;
	const pages = copy(await visioXml(pkg, pagesPart, 'Pages'));
	const priorCount = children(pages, 'Page').length;
	const rels = copy(await pkg.readXml(relsPart, 'Relationships'));
	const types = copy(await pkg.readXml('[Content_Types].xml', 'Types'));
	const dirty = new Map<string, Element>();
	const paths = new Set([...parts.keys()].map((path) => decodePath(path).toLowerCase()));
	// An unused content-type override also reserves its part name.
	for (const node of Array.from(types.childNodes)) {
		if (node.nodeType !== 1) continue;
		const name = (node as Element).getAttribute('PartName');
		if (name?.startsWith('/')) paths.add(decodePath(name.slice(1)).toLowerCase());
	}
	const relIds = new Set(
		Array.from(rels.childNodes)
			.filter((node) => node.nodeType === 1)
			.map((node) => (node as Element).getAttribute('Id')),
	);
	let nextPart = 1,
		nextRel = 1;
	for (const command of commands) {
		check();
		const existing = children(pages, 'Page');
		if (existing.some((page) => attribute(page, 'ID') === command.pageId))
			fail('EDIT_DUPLICATE_PAGE', 'New page ID already exists.');
		if (
			existing.some((page) =>
				['Name', 'NameU'].some(
					(name) => attribute(page, name)?.toLowerCase() === command.name.toLowerCase(),
				),
			)
		)
			fail('EDIT_DUPLICATE_PAGE_NAME', 'New page name already exists.');
		const after = existing.find((page) => attribute(page, 'ID') === command.afterPageId);
		if (!after) fail('EDIT_TARGET_NOT_FOUND', 'Insertion target page does not exist.');
		let path: string;
		do {
			check();
			path = `${directory}page${nextPart++}.xml`;
		} while (paths.has(path.toLowerCase()));
		paths.add(path.toLowerCase());
		let relId: string;
		do {
			check();
			relId = `rIdPage${nextRel++}`;
		} while (relIds.has(relId));
		relIds.add(relId);
		const doc = pages.ownerDocument!;
		const page = doc.createElementNS(pages.namespaceURI, 'Page');
		page.setAttribute('ID', command.pageId);
		page.setAttribute('Name', command.name);
		page.setAttribute('NameU', command.name);
		page.setAttribute('IsCustomName', '1');
		page.setAttribute('IsCustomNameU', '1');
		const sheet = child(after, 'PageSheet');
		if (sheet) page.appendChild(sheet.cloneNode(true));
		const background = attribute(after, 'BackPage');
		if (background !== undefined) page.setAttribute('BackPage', background);
		const link = doc.createElementNS(pages.namespaceURI, 'Rel');
		link.setAttributeNS(officeRel, 'r:id', relId);
		page.appendChild(link);
		pages.insertBefore(page, after.nextSibling);
		const relationship = rels.ownerDocument!.createElementNS(rels.namespaceURI, 'Relationship');
		relationship.setAttribute('Id', relId);
		relationship.setAttribute('Type', pageType);
		relationship.setAttribute('Target', path.slice(directory.length));
		rels.appendChild(relationship);
		const override = types.ownerDocument!.createElementNS(types.namespaceURI, 'Override');
		override.setAttribute('PartName', `/${path}`);
		override.setAttribute('ContentType', contentType);
		types.appendChild(override);
		const contentDoc = parseXml(`<PageContents xmlns="${pages.namespaceURI}"/>`);
		contentDoc.documentElement.setAttributeNS(
			'http://www.w3.org/XML/1998/namespace',
			'xml:space',
			'preserve',
		);
		dirty.set(path, contentDoc.documentElement);
	}
	dirty.set(pagesPart, pages);
	dirty.set(relsPart, rels);
	dirty.set('[Content_Types].xml', types);
	await updatePageAppProperties(pkg, pages, priorCount, dirty, limits, check);
	if (parts.size + commands.length > limits.maxEntries)
		fail('LIMIT_ENTRIES', 'Page insertion exceeds package entry limit.');
	let total = [...parts.values()].reduce((sum, bytes) => sum + bytes.length, 0),
		nodes = 0;
	for (const [path, root] of dirty) {
		const serialized = serializeEditedXml(root, limits, check);
		nodes += serialized.nodes;
		if (nodes > limits.maxTotalXmlNodes)
			fail('LIMIT_XML_TOTAL', 'Edited XML node total exceeds limit.');
		total += serialized.bytes.length - (parts.get(path)?.length ?? 0);
		if (total > limits.maxTotalBytes) fail('LIMIT_TOTAL', 'Edited package total exceeds limit.');
		parts.set(path, serialized.bytes);
	}
	const bytes = await writeEditedPackage(parts, maxOutput, deadline, check);
	const verified = await openEditablePackage(
		bytes,
		{ ...limits, maxInputBytes: maxOutput, maxRuntimeMs: Math.max(1, deadline - Date.now()) },
		check,
	);
	for (const command of commands)
		await visioXml(verified.pkg, verified.pages.get(command.pageId)!, 'PageContents');
	check();
	return {
		bytes,
		changedParts: [...dirty.keys()],
		diagnostics: [
			{
				code: 'edit-pages-experimental',
				message:
					'Blank pages were inserted with copied page settings. Page-count/index-dependent formula caches were not recalculated.',
			},
		],
	};
}
