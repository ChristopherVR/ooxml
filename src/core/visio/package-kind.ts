import JSZip from 'jszip';
import { buildXml, elements, parseXml } from '../xml/index';
import { relationshipsPartFor, resolvePartPath } from '../opc/relationships';
import {
	inspectPackageSignatures,
	isSignaturePart,
	stripPackageSignatures,
} from '../opc/signature/index';
import { VisioPackage, VisioPackageError } from './package';
import { decodePath, type VisioPackageLimits } from './package-common';

/** Main-part content types of the Visio package kinds (.vsdx, .vsdm, .vstx, .vstm, .vssx, .vssm). */
export const DRAWING_MAIN = 'application/vnd.ms-visio.drawing.main+xml';
export const DRAWING_MACRO_MAIN = 'application/vnd.ms-visio.drawing.macroEnabled.main+xml';
export const TEMPLATE_MAIN = 'application/vnd.ms-visio.template.main+xml';
export const TEMPLATE_MACRO_MAIN = 'application/vnd.ms-visio.template.macroEnabled.main+xml';
export const STENCIL_MAIN = 'application/vnd.ms-visio.stencil.main+xml';
export const STENCIL_MACRO_MAIN = 'application/vnd.ms-visio.stencil.macroEnabled.main+xml';

/** Types a package may carry when it is opened as a drawing (templates open as untitled drawings). */
export const DRAWING_MAIN_TYPES: ReadonlySet<string> = new Set([
	DRAWING_MAIN,
	DRAWING_MACRO_MAIN,
	TEMPLATE_MAIN,
	TEMPLATE_MACRO_MAIN,
]);
/** Types that can be read for their masters only. */
export const STENCIL_MAIN_TYPES: ReadonlySet<string> = new Set([
	STENCIL_MAIN,
	STENCIL_MACRO_MAIN,
	...DRAWING_MAIN_TYPES,
]);
/** Types an edit accepts; templates are converted to drawings first. */
export const EDITABLE_MAIN_TYPES: ReadonlySet<string> = new Set([DRAWING_MAIN, DRAWING_MACRO_MAIN]);

export const TEMPLATE_TO_DRAWING: ReadonlyMap<string, string> = new Map([
	[TEMPLATE_MAIN, DRAWING_MAIN],
	[TEMPLATE_MACRO_MAIN, DRAWING_MACRO_MAIN],
]);

export const DRAWING_TO_TEMPLATE: ReadonlyMap<string, string> = new Map([
	[DRAWING_MAIN, TEMPLATE_MAIN],
	[DRAWING_MACRO_MAIN, TEMPLATE_MACRO_MAIN],
]);

export interface VisioPackagePreparation {
	bytes: Uint8Array;
	/** The digital signatures were removed, since an edit invalidates them. */
	signatureRemoved: boolean;
	/** The template main content type was changed to the drawing one. */
	templateConverted: boolean;
}

const DIGITAL_SIGNATURE_TYPE = /digital-signature/i;
const RELS_PATH = /^_rels\/\.rels$|(^|\/)_rels\/[^/]+\.rels$/;

function relsSource(relsPath: string): string {
	return relsPath === '_rels/.rels' ? '' : relsPath.replace(/(^|\/)_rels\/([^/]+)\.rels$/, '$1$2');
}

function tryResolveTarget(source: string, target: string): string | undefined {
	try {
		return resolvePartPath(source, decodePath(target));
	} catch {
		return undefined;
	}
}

function decodedZipEntries(zip: JSZip): Map<string, string> {
	const map = new Map<string, string>();
	for (const name of Object.keys(zip.files)) {
		if (zip.files[name]?.dir) continue;
		try {
			map.set(decodePath(name), name);
		} catch {
			map.set(name, name);
		}
	}
	return map;
}

/**
 * Strip digital signatures from an already-bounded `JSZip`: both standard `_xmlsignatures/` and
 * non-standard locations (custom origin/signature/certificate folders via `opc/signature`, percent-
 * encoded signature paths, and signature relationships in part `.rels` files).
 */
export async function stripVisioSignatures(zip: JSZip): Promise<boolean> {
	const inspected = await inspectPackageSignatures(zip);
	if (inspected.hasSignatures) await stripPackageSignatures(zip);
	let removed = inspected.hasSignatures;
	const removedParts = new Set<string>(inspected.parts);

	for (const path of Object.keys(zip.files)) {
		if (zip.files[path]?.dir) continue;
		const decoded = decodePath(path);
		if (isSignaturePart(decoded)) {
			zip.remove(path);
			removedParts.add(path);
			removedParts.add(decoded);
			removed = true;
		}
	}

	// Sweep all remaining `.rels` files (including part `.rels` outside `_rels/.rels`) for any
	// digital-signature relationships or dangling references to already-removed signature parts.
	const keptTargets = new Set<string>();
	const candidateTargets = new Set<string>();
	for (const path of Object.keys(zip.files)) {
		const file = zip.files[path];
		if (!file || file.dir || !RELS_PATH.test(path)) continue;
		const source = relsSource(path);
		const doc = parseXml(await file.async('string'));
		let changed = false;
		for (const node of elements(doc.documentElement)) {
			if (node.localName !== 'Relationship') continue;
			const type = node.getAttribute('Type') ?? '';
			const target = node.getAttribute('Target');
			const external = node.getAttribute('TargetMode') === 'External';
			const resolved = !external && target !== null ? tryResolveTarget(source, target) : undefined;
			const isSigRel =
				DIGITAL_SIGNATURE_TYPE.test(type) ||
				(resolved !== undefined && (isSignaturePart(resolved) || removedParts.has(resolved)));
			if (isSigRel) {
				doc.documentElement.removeChild(node);
				changed = true;
				removed = true;
				if (resolved !== undefined) candidateTargets.add(resolved);
			} else if (resolved !== undefined) {
				keptTargets.add(resolved);
			}
		}
		if (changed) zip.file(path, buildXml(doc));
	}

	// Remove unreferenced signature target parts and transitively follow their own `.rels`.
	const byDecoded = decodedZipEntries(zip);
	for (const target of candidateTargets) {
		if (keptTargets.has(target)) continue;
		const actual = byDecoded.get(target);
		if (actual) {
			zip.remove(actual);
			removedParts.add(actual);
			removedParts.add(target);
			removed = true;
		}
		const relsPath = relationshipsPartFor(target);
		const actualRels = byDecoded.get(relsPath);
		const relsFile = actualRels ? zip.file(actualRels) : null;
		if (relsFile && actualRels) {
			const doc = parseXml(await relsFile.async('string'));
			for (const node of elements(doc.documentElement)) {
				if (node.localName !== 'Relationship') continue;
				const nextTarget = node.getAttribute('Target');
				if (node.getAttribute('TargetMode') === 'External' || nextTarget === null) continue;
				const next = tryResolveTarget(target, nextTarget);
				if (next !== undefined && !keptTargets.has(next)) candidateTargets.add(next);
			}
			zip.remove(actualRels);
			removedParts.add(actualRels);
			removedParts.add(relsPath);
			removed = true;
		}
	}

	// Remove any remaining signature defaults or overrides from `[Content_Types].xml`.
	const typesFile = zip.file('[Content_Types].xml');
	if (typesFile) {
		const doc = parseXml(await typesFile.async('string'));
		let changed = false;
		for (const node of elements(doc.documentElement)) {
			if (node.localName !== 'Default' && node.localName !== 'Override') continue;
			const contentType = node.getAttribute('ContentType') ?? '';
			const part = node.getAttribute('PartName');
			const resolved = part !== null ? tryResolveTarget('', part) : undefined;
			const isSigType =
				DIGITAL_SIGNATURE_TYPE.test(contentType) ||
				(resolved !== undefined && (isSignaturePart(resolved) || removedParts.has(resolved)));
			if (isSigType) {
				doc.documentElement.removeChild(node);
				changed = true;
				removed = true;
			}
		}
		if (changed) zip.file('[Content_Types].xml', buildXml(doc));
	}

	return removed;
}

/**
 * Make a package ready to be saved after an edit, as Visio does: digital signatures (whether in
 * the standard `_xmlsignatures/` layout or in non-standard parts/relationships) are dropped (an
 * edit invalidates them), and a template becomes a drawing of the same kind (a macro-enabled
 * template becomes a macro-enabled drawing). The VBA project, its relationship and the
 * macro-enabled content type are never touched, and macros are never run. Packages that need
 * neither change are returned as they are, byte for byte. Every part is read through the bounded
 * `VisioPackage` reader, so the input limits (size, entry count, compression ratio, deadline)
 * apply before anything is decompressed.
 */
export async function prepareVisioPackageForEdit(
	input: Uint8Array,
	limits: Partial<VisioPackageLimits> = {},
): Promise<VisioPackagePreparation> {
	const unchanged = { bytes: input, signatureRemoved: false, templateConverted: false };
	let pkg: VisioPackage;
	try {
		pkg = await VisioPackage.open(input, limits);
	} catch (error) {
		// A limit is final; anything else is reported by the edit's own open, with its own code.
		if (error instanceof VisioPackageError && error.code.startsWith('LIMIT_')) throw error;
		return unchanged;
	}
	const paths = pkg.paths();
	const decoder = new TextDecoder();
	let needsWork = paths.some((name) => isSignaturePart(decodePath(name)));
	if (!needsWork && paths.includes('[Content_Types].xml')) {
		const typesText = decoder.decode(await pkg.readBytes('[Content_Types].xml'));
		needsWork =
			DIGITAL_SIGNATURE_TYPE.test(typesText) ||
			typesText.includes(TEMPLATE_MAIN) ||
			typesText.includes(TEMPLATE_MACRO_MAIN);
	}
	if (!needsWork) {
		for (const path of paths) {
			if (!RELS_PATH.test(path)) continue;
			const relsText = decoder.decode(await pkg.readBytes(path));
			if (DIGITAL_SIGNATURE_TYPE.test(relsText)) {
				needsWork = true;
				break;
			}
		}
	}
	if (!needsWork) return unchanged;

	const zip = new JSZip();
	for (const path of paths) {
		zip.file(path, await pkg.readBytes(path), { createFolders: false });
	}
	let signatureRemoved = false;
	let templateConverted = false;
	try {
		signatureRemoved = await stripVisioSignatures(zip);
		const typesFile = zip.file('[Content_Types].xml');
		if (typesFile) {
			const doc = parseXml(await typesFile.async('string'));
			let changed = false;
			for (const node of elements(doc.documentElement)) {
				if (node.localName !== 'Override') continue;
				const contentType = node.getAttribute('ContentType') ?? '';
				const converted = TEMPLATE_TO_DRAWING.get(contentType);
				if (converted) {
					node.setAttribute('ContentType', converted);
					changed = true;
					templateConverted = true;
				}
			}
			if (changed) zip.file('[Content_Types].xml', buildXml(doc));
		}
	} catch (error) {
		if (error instanceof VisioPackageError && error.code.startsWith('LIMIT_')) throw error;
		return unchanged;
	}
	if (!signatureRemoved && !templateConverted) return unchanged;
	const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'STORE' });
	return { bytes, signatureRemoved, templateConverted };
}
