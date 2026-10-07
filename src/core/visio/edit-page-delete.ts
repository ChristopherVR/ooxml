import { attribute, child, children, yes } from './sheet';
import { related, visioXml } from './parts';
import { fail, decodePath } from './package-common';
import type { VisioPackage } from './package';
import { unquotedVisioFormula } from './formula-source';
import { editableVisioPageRoots } from './edit-page-roots';
import {
	relationshipsPartFor as pageRelationshipsPath,
	nextRelationshipId,
} from '../opc/relationships';
import { parseXml } from '../xml/index';

const officeRel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const pageType = 'http://schemas.microsoft.com/visio/2010/relationships/page';
const escape = (value: string) => value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/** Delete(0) freezes static references and clears backgrounds, retaining cached values. */
export async function deleteVisioPage(
	pkg: VisioPackage,
	pagesPart: string,
	pages: Element,
	rels: Element,
	types: Element,
	pagePaths: Map<string, string>,
	parts: Map<string, Uint8Array>,
	dirty: Map<string, Element>,
	removed: Set<string>,
	pageId: string,
	check: () => void,
): Promise<void> {
	const existing = children(pages, 'Page');
	const page = existing.find((node) => attribute(node, 'ID') === pageId);
	if (!page) fail('EDIT_TARGET_NOT_FOUND', 'Page does not exist.');
	const path = pagePaths.get(pageId)!;
	const names = new Set(
		['Name', 'NameU']
			.map((key) => attribute(page, key))
			.filter((value): value is string => !!value),
	);
	const reference = new RegExp(`Pages\\[(?:${[...names].map(escape).join('|')})\\]!`, 'i');
	const documentPart = (await related(pkg, '', 'document'))!;
	const roots = await editableVisioPageRoots(
		pkg,
		pagesPart,
		pages,
		pagePaths,
		dirty,
		check,
		pageId,
	);
	for (const [part, root] of roots) {
		check();
		let changed = false;
		for (const node of Array.from(root.getElementsByTagName('*'))) {
			check();
			// The removed page's own sheet is discarded, not evaluated.
			let ancestor: Node | null = node;
			while (ancestor && (ancestor as Element).localName !== 'Page') ancestor = ancestor.parentNode;
			if (ancestor === page) continue;
			const formula = attribute(node, 'F');
			if (!formula) continue;
			const code = unquotedVisioFormula(formula);
			if (
				/\b(?:INDIRECT|EVALCELL|EVALTEXT|GETREF|REF|SETATREF|SETATREFEXPR|SETATREFEVAL|SHAPETEXT|CALL|RUNADDON)\s*\(/i.test(
					code,
				)
			)
				fail(
					'EDIT_UNSUPPORTED_PAGE_DELETE_DEPENDENCY',
					'Dynamic page deletion dependencies cannot be safely frozen.',
				);
			if (!reference.test(code)) continue;
			if (node.localName !== 'Cell' || !node.hasAttribute('V') || node.hasAttribute('E'))
				fail(
					'EDIT_UNSUPPORTED_PAGE_DELETE_CACHE',
					'Referenced page deletion requires a valid local cell cache.',
				);
			node.removeAttribute('F');
			changed = true;
		}
		if (changed) dirty.set(part, root);
	}
	for (const candidate of existing)
		if (attribute(candidate, 'BackPage') === pageId) candidate.removeAttribute('BackPage');
	for (const relationshipPart of pkg.paths()) {
		check();
		if (
			!relationshipPart.endsWith('.rels') ||
			relationshipPart === pageRelationshipsPath(pagesPart) ||
			relationshipPart === pageRelationshipsPath(path) ||
			!parts.has(relationshipPart)
		)
			continue;
		const sourcePart =
			relationshipPart === '_rels/.rels'
				? ''
				: relationshipPart.replace(/(^|\/)_rels\/([^/]+)\.rels$/, '$1$2');
		if (sourcePart === relationshipPart) continue;
		const incoming = [...(await pkg.relationships(sourcePart)).values()].filter(
			(rel) => rel.mode === 'Internal' && rel.target === path,
		);
		if (!incoming.length) continue;
		if (incoming.some((rel) => rel.type !== pageType))
			fail(
				'EDIT_UNSUPPORTED_PAGE_DELETE_RELATIONSHIP',
				'Deleted page has an unsupported incoming relationship.',
			);
		if (!sourcePart)
			fail(
				'EDIT_UNSUPPORTED_PAGE_DELETE_RELATIONSHIP',
				'Cannot remove a package-root page reference.',
			);
		const sourceRoot =
			dirty.get(sourcePart) ?? roots.get(sourcePart) ?? (await pkg.readXml(sourcePart));
		if (
			Array.from(sourceRoot.getElementsByTagName('*')).some((node) =>
				incoming.some((rel) => node.getAttributeNS(officeRel, 'id') === rel.id),
			)
		)
			fail(
				'EDIT_UNSUPPORTED_PAGE_DELETE_RELATIONSHIP',
				'Page relationship is used by unsupported explicit XML markup.',
			);
		const source =
			dirty.get(relationshipPart) ?? (await pkg.readXml(relationshipPart, 'Relationships'));
		const root = (source.ownerDocument!.cloneNode(true) as Document).documentElement;
		for (const node of Array.from(root.childNodes))
			if (
				node.nodeType === 1 &&
				incoming.some((rel) => rel.id === (node as Element).getAttribute('Id'))
			)
				root.removeChild(node);
		if (Array.from(root.childNodes).some((node) => node.nodeType === 1))
			dirty.set(relationshipPart, root);
		else {
			parts.delete(relationshipPart);
			dirty.delete(relationshipPart);
			removed.add(relationshipPart);
		}
	}
	const relId = child(page, 'Rel')?.getAttributeNS(officeRel, 'id');
	const relationship = Array.from(rels.childNodes).find(
		(node) => node.nodeType === 1 && (node as Element).getAttribute('Id') === relId,
	);
	if (!relationship) fail('INVALID_RELATIONSHIP', 'Deleted page relationship is missing.');
	rels.removeChild(relationship);
	pages.removeChild(page);
	pagePaths.delete(pageId);
	for (const deleted of [path, pageRelationshipsPath(path)]) {
		if (parts.delete(deleted) || dirty.delete(deleted)) removed.add(deleted);
		dirty.delete(deleted);
	}
	for (const node of Array.from(types.childNodes)) {
		if (node.nodeType !== 1) continue;
		const name = (node as Element).getAttribute('PartName');
		if (
			name &&
			[path, pageRelationshipsPath(path)].some(
				(part) => decodePath(name.slice(1)).toLowerCase() === decodePath(part).toLowerCase(),
			)
		)
			types.removeChild(node);
	}
	if (!children(pages, 'Page').some((page) => !yes(attribute(page, 'Background'))))
		addReplacementPage(pagesPart, pages, rels, types, pagePaths, parts, dirty, existing, check);
	// Restore window page identity to the first surviving foreground page.
	const windowsPart = await related(pkg, documentPart, 'windows', false);
	if (windowsPart) {
		const original = dirty.get(windowsPart) ?? (await visioXml(pkg, windowsPart, 'Windows'));
		const root = (original.ownerDocument!.cloneNode(true) as Document).documentElement;
		let changed = false;
		for (const window of children(root, 'Window'))
			if (attribute(window, 'Page') === pageId) {
				window.setAttribute(
					'Page',
					attribute(
						children(pages, 'Page').find((node) => !yes(attribute(node, 'Background'))),
						'ID',
					)!,
				);
				changed = true;
			}
		if (changed) dirty.set(windowsPart, root);
	}
	dirty.set(pagesPart, pages);
	dirty.set(pageRelationshipsPath(pagesPart), rels);
	dirty.set('[Content_Types].xml', types);
}

/** Native last-page deletion leaves a new blank foreground page with default page settings. */
function addReplacementPage(
	pagesPart: string,
	pages: Element,
	rels: Element,
	types: Element,
	pagePaths: Map<string, string>,
	parts: Map<string, Uint8Array>,
	dirty: Map<string, Element>,
	previous: Element[],
	check: () => void,
): void {
	let id = 0;
	for (const page of previous) {
		const value = Number(attribute(page, 'ID'));
		if (!Number.isSafeInteger(value) || value < 0 || value >= 4294967295)
			fail('EDIT_UNSUPPORTED_PAGE_ID', 'Cannot allocate replacement page ID.');
		id = Math.max(id, value + 1);
	}
	const names = new Set(
		previous.flatMap((node) => ['Name', 'NameU'].map((key) => attribute(node, key)?.toLowerCase())),
	);
	let nameIndex = previous.length + 1;
	while (names.has(`page-${nameIndex}`)) {
		check();
		nameIndex++;
	}
	const directory = pagesPart.slice(0, pagesPart.lastIndexOf('/') + 1);
	const reserved = new Set(
		[...parts.keys(), ...dirty.keys()].map((path) => decodePath(path).toLowerCase()),
	);
	for (const node of Array.from(types.childNodes)) {
		if (node.nodeType !== 1) continue;
		const name = (node as Element).getAttribute('PartName');
		if (name) reserved.add(decodePath(name.slice(1)).toLowerCase());
	}
	let number = 1,
		path: string;
	do {
		check();
		path = `${directory}page${number++}.xml`;
	} while (reserved.has(path.toLowerCase()));
	const relIds = new Set(
		Array.from(rels.childNodes)
			.filter((node) => node.nodeType === 1)
			.map((node) => (node as Element).getAttribute('Id') ?? ''),
	);
	const relId = nextRelationshipId(relIds);
	const page = pages.ownerDocument!.createElementNS(pages.namespaceURI, 'Page');
	page.setAttribute('ID', String(id));
	page.setAttribute('Name', `Page-${nameIndex}`);
	page.setAttribute('NameU', `Page-${nameIndex}`);
	const sheet = pages.ownerDocument!.createElementNS(pages.namespaceURI, 'PageSheet');
	for (const [name, value] of [
		['PageWidth', '8.5'],
		['PageHeight', '11'],
		['PageScale', '1'],
		['DrawingScale', '1'],
	]) {
		const cell = pages.ownerDocument!.createElementNS(pages.namespaceURI, 'Cell');
		cell.setAttribute('N', name!);
		cell.setAttribute('V', value!);
		sheet.appendChild(cell);
	}
	page.appendChild(sheet);
	const link = pages.ownerDocument!.createElementNS(pages.namespaceURI, 'Rel');
	link.setAttributeNS(officeRel, 'r:id', relId);
	page.appendChild(link);
	pages.appendChild(page);
	const relationship = rels.ownerDocument!.createElementNS(rels.namespaceURI, 'Relationship');
	relationship.setAttribute('Id', relId);
	relationship.setAttribute('Type', pageType);
	relationship.setAttribute('Target', path.slice(directory.length));
	rels.appendChild(relationship);
	const override = types.ownerDocument!.createElementNS(types.namespaceURI, 'Override');
	override.setAttribute('PartName', `/${path}`);
	override.setAttribute('ContentType', 'application/vnd.ms-visio.page+xml');
	types.appendChild(override);
	pagePaths.set(String(id), path);
	dirty.set(
		path,
		parseXml(`<PageContents xmlns="${pages.namespaceURI}" xml:space="preserve"/>`).documentElement,
	);
}
