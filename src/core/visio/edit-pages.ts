import { attribute, child, children } from './sheet';
import { indexedPart, related, visioXml } from './parts';
import { VisioPackage } from './package';
import { decodePath, fail, type VisioPackageLimits } from './package-common';
import { openEditablePackage, writeEditedPackage } from './edit-package';
import { serializeEditedXml } from './edit-text';
import { parseXml } from '../xml/index';
import { updatePageAppProperties } from './edit-page-properties';
import { recalculatePageFormulas } from './edit-page-formulas';
import { renameVisioPage } from './edit-page-rename';
import { deleteVisioPage } from './edit-page-delete';
import { relationshipsPartFor } from '../opc/relationships';
import type { VisioPageEdit } from './edit-commands';
import type { EditVsdxResult } from './edit';

const officeRel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const pageType = 'http://schemas.microsoft.com/visio/2010/relationships/page';
const contentType = 'application/vnd.ms-visio.page+xml';
function copy(root: Element): Element {
	return (root.ownerDocument!.cloneNode(true) as Document).documentElement;
}

/** Source-backed page metadata/OPC transaction with bounded local numeric cache refresh. */
export async function editVsdxPages(
	original: Uint8Array,
	pkg: VisioPackage,
	parts: Map<string, Uint8Array>,
	commands: readonly VisioPageEdit[],
	limits: VisioPackageLimits,
	maxOutput: number,
	deadline: number,
	check: () => void,
): Promise<EditVsdxResult> {
	const documentPart = (await related(pkg, '', 'document'))!;
	const pagesPart = (await related(pkg, documentPart, 'pages'))!;
	const slash = pagesPart.lastIndexOf('/');
	const directory = pagesPart.slice(0, slash + 1);
	const relsPart = relationshipsPartFor(pagesPart);
	const pages = copy(await visioXml(pkg, pagesPart, 'Pages'));
	const priorCount = children(pages, 'Page').length;
	const pagePaths = new Map<string, string>();
	for (const page of children(pages, 'Page'))
		pagePaths.set(attribute(page, 'ID')!, await indexedPart(pkg, pagesPart, page, 'page'));
	const rels = copy(await pkg.readXml(relsPart, 'Relationships'));
	const types = copy(await pkg.readXml('[Content_Types].xml', 'Types'));
	const dirty = new Map<string, Element>();
	const removed = new Set<string>();
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
		if (command.type === 'delete-page') {
			await deleteVisioPage(
				pkg,
				pagesPart,
				pages,
				rels,
				types,
				pagePaths,
				parts,
				dirty,
				removed,
				command.pageId,
				check,
			);
			continue;
		}
		if (command.type === 'rename-page') {
			await renameVisioPage(
				pkg,
				pagesPart,
				pages,
				pagePaths,
				dirty,
				command.pageId,
				command.name,
				check,
			);
			continue;
		}
		if (command.type === 'reorder-page') {
			const target = existing.find((page) => attribute(page, 'ID') === command.pageId);
			if (!target) fail('EDIT_TARGET_NOT_FOUND', 'Page does not exist.');
			if (command.index >= existing.length)
				fail('INVALID_EDIT', 'Page index exceeds the page list.');
			if (existing.indexOf(target) === command.index) continue;
			const remaining = existing.filter((page) => page !== target);
			pages.insertBefore(target, remaining[command.index] ?? null);
			dirty.set(pagesPart, pages);
			continue;
		}
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
		pagePaths.set(command.pageId, path);
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
		dirty.set(pagesPart, pages);
		dirty.set(relsPart, rels);
		dirty.set('[Content_Types].xml', types);
	}
	if (!dirty.size) {
		if (original.length > maxOutput)
			fail('LIMIT_EDIT_OUTPUT', 'Saved package exceeds output limit.');
		return { bytes: original, changedParts: [], diagnostics: [] };
	}
	await updatePageAppProperties(pkg, pages, priorCount, dirty, limits, check);
	await recalculatePageFormulas(pkg, pagesPart, pages, pagePaths, dirty, check);
	if (new Set([...parts.keys(), ...dirty.keys()]).size > limits.maxEntries)
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
	const pageParts = new Set(verified.pages.values());
	for (const path of dirty.keys())
		if (pageParts.has(path)) await visioXml(verified.pkg, path, 'PageContents');
	check();
	return {
		bytes,
		changedParts: [...new Set([...removed, ...dirty.keys()])],
		diagnostics: [
			{
				code: 'edit-pages-experimental',
				message:
					'Page metadata and supported local numeric page-dependent caches were updated. Inherited, string and background-render-context formulas remain unsupported.',
			},
		],
	};
}
