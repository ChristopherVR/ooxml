import JSZip from 'jszip';
import { decodePath } from './package-common';
import { isSignaturePart } from '../opc/signature/index';

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

const TEMPLATE_TO_DRAWING = new Map([
	[TEMPLATE_MAIN, DRAWING_MAIN],
	[TEMPLATE_MACRO_MAIN, DRAWING_MACRO_MAIN],
]);

export interface VisioPackagePreparation {
	bytes: Uint8Array;
	/** The digital signatures were removed, since an edit invalidates them. */
	signatureRemoved: boolean;
	/** The template main content type was changed to the drawing one. */
	templateConverted: boolean;
}

/** The value of an attribute of one start tag, quoted either way. */
const attr = (tag: string, name: string): string | undefined =>
	new RegExp(String.raw`\b${name}\s*=\s*(?:"([^"]*)"|'([^']*)')`)
		.exec(tag)
		?.slice(1)
		.find((value) => value !== undefined);
const DIGITAL_SIGNATURE_TYPE = /digital-signature/i;

/**
 * Make a package ready to be saved after an edit, as Visio does: the digital signatures are
 * dropped (an edit invalidates them), and a template becomes a drawing of the same kind (a
 * macro-enabled template becomes a macro-enabled drawing). The VBA project, its relationship and
 * the macro-enabled content type are never touched, and macros are never run. Packages that need
 * neither change are returned as they are, byte for byte. Input that is not a readable ZIP is
 * returned unchanged for the caller's own validation to refuse.
 */
export async function prepareVisioPackageForEdit(
	input: Uint8Array,
): Promise<VisioPackagePreparation> {
	const unchanged = { bytes: input, signatureRemoved: false, templateConverted: false };
	let zip: JSZip;
	try {
		zip = await JSZip.loadAsync(input);
	} catch {
		return unchanged;
	}
	const signed = Object.keys(zip.files).filter((name) => isSignaturePart(decodePath(name)));
	const typesFile = zip.file('[Content_Types].xml');
	const relsFile = zip.file('_rels/.rels');
	let types = typesFile ? await typesFile.async('string') : '';
	let rels = relsFile ? await relsFile.async('string') : '';
	let signatureRemoved = signed.length > 0;
	let templateConverted = false;
	const newRels = rels.replace(
		/<Relationship\b[^>]*\/>|<Relationship\b[^>]*>[\s\S]*?<\/Relationship>/g,
		(tag) => {
			if (!DIGITAL_SIGNATURE_TYPE.test(attr(tag, 'Type') ?? '')) return tag;
			signatureRemoved = true;
			return '';
		},
	);
	const newTypes = types.replace(/<(Override|Default)\b[^>]*\/>/g, (tag) => {
		const contentType = attr(tag, 'ContentType') ?? '';
		const part = attr(tag, 'PartName');
		if (DIGITAL_SIGNATURE_TYPE.test(contentType) || (part && isSignaturePart(part))) {
			signatureRemoved = true;
			return '';
		}
		const converted = TEMPLATE_TO_DRAWING.get(contentType);
		if (converted && tag.startsWith('<Override')) {
			templateConverted = true;
			return tag.replace(contentType, converted);
		}
		return tag;
	});
	if (!signatureRemoved && !templateConverted) return unchanged;
	types = newTypes;
	rels = newRels;
	for (const name of signed) zip.remove(name);
	if (typesFile) zip.file('[Content_Types].xml', types);
	if (relsFile) zip.file('_rels/.rels', rels);
	// Every other part is stored as it was read; a ZIP is rewritten only when it had to change.
	const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
	return { bytes, signatureRemoved, templateConverted };
}
