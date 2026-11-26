import { createOfficeGuid } from '../crypto/uuid';
import { relationshipsPartFor } from '../opc/relationships';
import { elements, NS } from '../xml/index';
import { createVisioPackage } from './create-document';
import { addMasterToDrawing, copy, shapesOf } from './edit-master-add';
import { importStencilMaster } from './edit-stencil-import';
import { assertShapeLocks } from './edit-style-admission';
import { fail } from './package-common';
import { decodeVisioPlainText, encodeVisioPlainText } from './plain-text';
import { indexedPart, related, visioXml } from './parts';
import { attribute, child, children } from './sheet';
import { addSelectedStencilShape } from './stencil-from-shape';
import { ensureStencilIcon } from './stencil-icon';
import { stencilMasterContents, stencilMasterEntry } from './stencil-master-part';
import { stencilMasterGuid, visioBuiltInMaster } from './stencil-masters';
import { removeStencilMasterParts } from './stencil-remove-parts';
import {
	openStencil,
	saveStencilParts,
	stencilSession,
	type VisioStencilEditOptions,
} from './stencil-edit-session';
export type { VisioStencilEditOptions };

export type VisioStencilEdit =
	| { type: 'add-built-in'; id: string }
	| { type: 'import-master'; stencil: Uint8Array; id: string }
	| { type: 'add-shape'; drawing: Uint8Array; pageId: string; shapeId: string; name: string }
	| { type: 'rename-master'; id: string; name: string }
	| { type: 'delete-master'; id: string }
	| { type: 'edit-master-text'; id: string; shapeId: string; text: string };

export interface VisioStencilSave {
	bytes: Uint8Array;
	extension: 'vssx' | 'vssm';
	signatureRemoved: boolean;
}

export function createVisioStencil(): Promise<Uint8Array> {
	return createVisioPackage({}, true);
}

export function stencilName(name: string): string {
	const trimmed = name.trim();
	if (!trimmed || trimmed.length > 255 || /[\u0000-\u001f\ufffe\uffff]/u.test(trimmed))
		fail('INVALID_MASTER_NAME', 'Enter a master name between 1 and 255 characters.');
	return trimmed;
}

/** Save the container with opaque parts intact. Package signatures are removed on an edit. */
export async function editVisioStencil(
	source: Uint8Array,
	edit?: VisioStencilEdit,
	options: VisioStencilEditOptions = {},
): Promise<VisioStencilSave> {
	const session = stencilSession(options);
	const { limits, deadline, check } = session;
	const workspace = await openStencil(source, session);
	const { pkg, parts, mastersPart, signatureRemoved, extension } = workspace;
	const finish = (bytes: Uint8Array): VisioStencilSave => ({ bytes, extension, signatureRemoved });
	if (!edit) return { bytes: source, extension, signatureRemoved: false };
	if (edit.type === 'add-built-in') {
		const builtIn = visioBuiltInMaster(edit.id);
		if (!builtIn) fail('EDIT_TARGET_NOT_FOUND', 'The stencil has no such master.');
		const result = await addMasterToDrawing(
			workspace.source,
			{
				uniqueId: builtIn.uniqueId,
				name: builtIn.master.name,
				entry(masters, id, name, relId) {
					ensureStencilIcon(stencilMasterEntry(masters, builtIn, id, name, relId), builtIn.master);
				},
				contents: (document) => stencilMasterContents(document.namespaceURI!, document, builtIn),
			},
			limits,
			limits.maxInputBytes,
			deadline,
			check,
			true,
		);
		return finish(result.bytes);
	}
	if (edit.type === 'import-master') {
		const result = await importStencilMaster(
			workspace.source,
			edit.stencil,
			edit.id,
			limits,
			limits.maxInputBytes,
			deadline,
			check,
			{ stencil: true, icon: ensureStencilIcon },
		);
		return finish(result.bytes);
	}
	if (edit.type === 'add-shape')
		return finish(
			await addSelectedStencilShape(
				workspace.source,
				{ ...edit, name: stencilName(edit.name) },
				session,
			),
		);
	if (!mastersPart) fail('EDIT_TARGET_NOT_FOUND', 'The stencil has no masters.');
	const masters = copy(await visioXml(pkg, mastersPart, 'Masters'));
	const entries = children(masters, 'Master');
	const entry = entries.find((node) => attribute(node, 'ID') === edit.id);
	if (!entry) fail('EDIT_TARGET_NOT_FOUND', 'The stencil has no such master.');
	const path = await indexedPart(pkg, mastersPart, entry, 'master');
	const dirty = new Map<string, Element>([[mastersPart, masters]]);
	if (edit.type === 'rename-master') {
		const name = stencilName(edit.name);
		if (attribute(entry, 'Name') === name && attribute(entry, 'NameU') === name)
			return finish(workspace.source);
		if (
			entries.some(
				(node) =>
					node !== entry &&
					['Name', 'NameU'].some(
						(key) => attribute(node, key)?.toLowerCase() === name.toLowerCase(),
					),
			)
		)
			fail('INVALID_MASTER_NAME', 'A master already has that name.');
		for (const key of ['Name', 'NameU']) entry.setAttribute(key, name);
		for (const key of ['IsCustomName', 'IsCustomNameU']) entry.setAttribute(key, '1');
	} else if (edit.type === 'delete-master') {
		const relPart = relationshipsPartFor(mastersPart);
		const rels = copy(await pkg.readXml(relPart, 'Relationships'));
		const relId = child(entry, 'Rel')?.getAttributeNS(NS.r, 'id');
		const otherEntries = [
			...entries.filter((node) => node !== entry),
			...children(masters, 'MasterShortcut'),
		];
		const relationships = await pkg.relationships(mastersPart);
		if (
			otherEntries.some((node) => {
				const id = child(node, 'Rel')?.getAttributeNS(NS.r, 'id') ?? '';
				return relationships.get(id)?.target === path;
			})
		)
			fail('UNSUPPORTED_EDIT_PACKAGE', 'Another master entry uses this master part.');
		for (const node of elements(rels))
			if (attribute(node, 'Id') === relId) node.parentNode?.removeChild(node);
		const typesPath = '[Content_Types].xml';
		const types = copy(await pkg.readXml(typesPath, 'Types'));
		await removeStencilMasterParts({
			pkg,
			parts,
			types,
			master: path,
			index: mastersPart,
			relId,
			check,
		});
		dirty.set(typesPath, types);
		masters.removeChild(entry);
		dirty.set(relPart, rels);
	} else {
		const contents = copy(await visioXml(pkg, path, 'MasterContents'));
		const shape = shapesOf(contents).find((node) => attribute(node, 'ID') === edit.shapeId);
		if (!shape || shape.hasAttribute('Master') || attribute(shape, 'Type') !== 'Shape')
			fail('UNSUPPORTED_TEXT_EDIT', 'Select a top-level shape with editable plain text.');
		const documentPart = (await related(pkg, '', 'document'))!;
		assertShapeLocks(shape, await visioXml(pkg, documentPart, 'VisioDocument'), ['LockTextEdit']);
		const text = child(shape, 'Text');
		if (
			children(shape, 'Text').length > 1 ||
			(text && elements(text).length) ||
			children(shape, 'Section').some((section) => attribute(section, 'N') === 'Field')
		)
			fail('UNSUPPORTED_TEXT_EDIT', 'Select a top-level shape with editable plain text.');
		if (
			edit.text.length > 4096 ||
			/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/u.test(edit.text)
		)
			fail('UNSUPPORTED_TEXT_EDIT', 'Enter plain text of up to 4096 characters.');
		const value = edit.text.replace(/\r\n?/g, '\n');
		if (decodeVisioPlainText(text?.textContent ?? '') === value) return finish(workspace.source);
		const doc = contents.ownerDocument!;
		const node = text ?? doc.createElementNS(contents.namespaceURI, 'Text');
		node.textContent = encodeVisioPlainText(value);
		if (!text) shape.appendChild(node);
		dirty.set(path, contents);
	}
	// Visio distinguishes revised copies by UniqueID and keeps the family's BaseID.
	if (edit.type !== 'delete-master') {
		if (!entry.hasAttribute('BaseID')) {
			entry.setAttribute('BaseID', attribute(entry, 'UniqueID') ?? createOfficeGuid());
		}
		entry.setAttribute('UniqueID', stencilMasterGuid(createOfficeGuid(), true));
	}
	return finish(await saveStencilParts(parts, dirty, session));
}
