import type { EditVsdxResult } from './edit';
import { copy, shapesOf } from './edit-master-add';
import { openEditablePackage, writeEditedPackage } from './edit-package';
import { serializeEditedXml } from './edit-text';
import {
	applyVisioMasterFormat,
	changesMasterSize,
	inheritsMasterSize,
	masterPlainText,
	refreshInheritedCaches,
	releasePlaceholderText,
	snapshotVisioMasterFormat,
	type VisioMasterFormat,
} from './master-format';
import type { VisioPackage } from './package';
import { fail, type VisioPackageLimits } from './package-common';
import { indexedPart, related, visioXml } from './parts';
import { attribute, children } from './sheet';

/**
 * Edit Master on a master of the drawing's Document Stencil: its default text, size, fill, line
 * and text formatting change in the master, and every instance on every page shows the change
 * where it inherits the cell (its own formatting stays, as in Visio). Changing the size of a
 * master that instances take their size from is refused for now: their pins, text blocks and
 * glued connectors would have to be laid out again.
 */
export interface VisioDocumentMasterEdit {
	type: 'edit-document-master';
	/** The page the edit is made from; the master belongs to the whole drawing. */
	pageId: string;
	masterId: string;
	format: VisioMasterFormat;
}

export const isVisioDocumentMasterEdit = (edit: {
	type: string;
}): edit is VisioDocumentMasterEdit => edit.type === 'edit-document-master';

export function snapshotDocumentMasterEdit(edit: VisioDocumentMasterEdit): VisioDocumentMasterEdit {
	if (typeof edit.pageId !== 'string' || !edit.pageId || edit.pageId.length > 256)
		fail('INVALID_EDIT', 'Invalid edit target.');
	if (typeof edit.masterId !== 'string' || !/^(0|[1-9]\d{0,9})$/.test(edit.masterId))
		fail('INVALID_EDIT', 'Master IDs must be canonical unsigned integers.');
	return {
		type: 'edit-document-master',
		pageId: edit.pageId,
		masterId: edit.masterId,
		format: snapshotVisioMasterFormat(edit.format),
	};
}

/** The instances of `masterId` on a page: top-level shapes that name it, at any depth. */
function instances(root: Element, masterId: string): Element[] {
	const found: Element[] = [];
	const pending = [root];
	while (pending.length) {
		const parent = pending.pop()!;
		for (const shape of shapesOf(parent)) {
			if (attribute(shape, 'Master') === masterId && !shape.hasAttribute('MasterShape'))
				found.push(shape);
			pending.push(shape);
		}
	}
	return found;
}

export async function editVsdxDocumentMaster(
	pkg: VisioPackage,
	parts: Map<string, Uint8Array>,
	pages: ReadonlyMap<string, string>,
	edit: VisioDocumentMasterEdit,
	limits: VisioPackageLimits,
	maxOutput: number,
	deadline: number,
	check: () => void,
): Promise<EditVsdxResult> {
	if (!pages.has(edit.pageId)) fail('EDIT_TARGET_NOT_FOUND', 'Page does not exist.');
	const documentPart = await related(pkg, '', 'document');
	if (!documentPart) fail('INVALID_DOCUMENT', 'Missing Visio document part.');
	const mastersPart = await related(pkg, documentPart!, 'masters', false);
	if (!mastersPart) fail('EDIT_TARGET_NOT_FOUND', 'The drawing has no masters.');
	const masters = await visioXml(pkg, mastersPart!, 'Masters');
	const entry = children(masters, 'Master').find((node) => attribute(node, 'ID') === edit.masterId);
	if (!entry) fail('EDIT_TARGET_NOT_FOUND', 'The drawing has no such master.');
	const masterPart = await indexedPart(pkg, mastersPart!, entry!, 'master');
	const contents = copy(await visioXml(pkg, masterPart, 'MasterContents'));
	const roots = shapesOf(contents).filter((shape) => attribute(shape, 'Del') !== '1');
	if (roots.length !== 1)
		fail('UNSUPPORTED_MASTER_EDIT', 'Edit Master changes a master with one top-level shape.');
	const document = await visioXml(pkg, documentPart!, 'VisioDocument');
	const releaseText = edit.format.text !== undefined && !masterPlainText(roots[0]!);
	const { changes, changed } = applyVisioMasterFormat(roots[0]!, document, edit.format);
	if (!changed)
		return {
			bytes: await writeEditedPackage(parts, maxOutput, deadline, check),
			changedParts: [],
			diagnostics: [],
		};
	const dirty = new Map<string, Element>([[masterPart, contents]]);
	for (const [, path] of pages) {
		check();
		const root = copy(await visioXml(pkg, path, 'PageContents'));
		let touched = false;
		for (const instance of instances(root, edit.masterId)) {
			if (changesMasterSize(edit.format) && inheritsMasterSize(instance))
				fail(
					'UNSUPPORTED_MASTER_EDIT',
					'Shapes on the drawing take their size from this master; resize them instead of the master.',
				);
			if (refreshInheritedCaches(instance, changes)) touched = true;
			if (releaseText && releasePlaceholderText(instance)) touched = true;
		}
		if (touched) dirty.set(path, root);
	}
	let nodes = 0;
	for (const [name, xml] of dirty) {
		const serialized = serializeEditedXml(xml, limits, check);
		nodes += serialized.nodes;
		if (nodes > limits.maxTotalXmlNodes)
			fail('LIMIT_XML_TOTAL', 'Edited XML node total exceeds limit.');
		parts.set(name, serialized.bytes);
	}
	const bytes = await writeEditedPackage(parts, maxOutput, deadline, check);
	const verified = await openEditablePackage(
		bytes,
		{ ...limits, maxInputBytes: maxOutput, maxRuntimeMs: Math.max(1, deadline - Date.now()) },
		check,
	);
	await visioXml(verified.pkg, masterPart, 'MasterContents');
	check();
	return { bytes, changedParts: [...dirty.keys()], diagnostics: [] };
}
