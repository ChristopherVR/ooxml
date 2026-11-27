import type { VisioMetafileTreeConverter } from './convert-metafile';
import type { VisioForeignVector } from './foreign-vector';
import { readVisioImagePart, type VisioImageOptions } from './media';
import type { VisioImage, VisioOleObject, VisioOleObjectFormat } from './model';
import { inspectVisioMetafilePart } from './prepare-metafiles';
import { VisioPackage } from './package';
import { fail } from './package-common';
import { children, type Report } from './sheet';
import { NS } from '../xml/index';

const PACKAGE = `${NS.r}/package`;
const OLE_OBJECT = `${NS.r}/oleObject`;
const IMAGE = `${NS.r}/image`;

/** Package content types, and the extension Visio and Office give each, by format. */
const FORMATS: readonly {
	format: VisioOleObjectFormat;
	types: readonly string[];
	extension: RegExp;
}[] = [
	{
		format: 'docx',
		types: ['application/vnd.openxmlformats-officedocument.wordprocessingml.document'],
		extension: /\.docx$/i,
	},
	{
		format: 'xlsx',
		types: ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'],
		extension: /\.xlsx$/i,
	},
	{
		format: 'pptx',
		types: ['application/vnd.openxmlformats-officedocument.presentationml.presentation'],
		extension: /\.pptx$/i,
	},
	{ format: 'vsdx', types: ['application/vnd.ms-visio.drawing'], extension: /\.vsdx$/i },
	{
		format: 'macro-enabled',
		types: [
			'application/vnd.ms-word.document.macroEnabled.12',
			'application/vnd.ms-excel.sheet.macroEnabled.12',
			'application/vnd.ms-powerpoint.presentation.macroEnabled.12',
			'application/vnd.ms-visio.drawing.macroEnabled.12',
		],
		extension: /\.(docm|xlsm|pptm|vsdm)$/i,
	},
	{
		format: 'ole',
		types: ['application/vnd.openxmlformats-officedocument.oleObject'],
		extension: /\.bin$/i,
	},
];

const REFUSALS: Partial<Record<VisioOleObjectFormat, string>> = {
	'macro-enabled':
		'This embedded document contains macros, so it is not opened here; it is kept as it is.',
	ole: 'This object is stored for the program that created it; it is shown from its saved picture and kept as it is, but cannot be opened here.',
	other:
		'This kind of embedded file is not opened here; it is shown from its saved picture and kept as it is.',
};

/** What an embedded part holds, from its content type first and its extension otherwise. */
export function classifyVisioOleEmbedding(
	part: string,
	relationshipType: string,
	contentType: string | undefined,
): VisioOleObjectFormat {
	const type = contentType?.toLowerCase();
	const byType = FORMATS.find((entry) => type && entry.types.some((t) => t.toLowerCase() === type));
	if (byType) return byType.format;
	// A declared content type that names none of these is trusted over the part's extension.
	if (type) return relationshipType === OLE_OBJECT ? 'ole' : 'other';
	return FORMATS.find((entry) => entry.extension.test(part))?.format ?? 'other';
}

/** The content type `[Content_Types].xml` declares for `part`: its Override, else its Default. */
async function contentTypeOf(pkg: VisioPackage, part: string): Promise<string | undefined> {
	if (!pkg.has('[Content_Types].xml')) return undefined;
	const types = await pkg.readXml('[Content_Types].xml', 'Types');
	let fallback: string | undefined;
	const extension = /\.([^./]+)$/.exec(part)?.[1]?.toLowerCase();
	for (const node of Array.from(types.children)) {
		if (
			node.localName === 'Override' &&
			node.getAttribute('PartName')?.toLowerCase() === `/${part}`.toLowerCase()
		)
			return node.getAttribute('ContentType') ?? undefined;
		if (node.localName === 'Default' && node.getAttribute('Extension')?.toLowerCase() === extension)
			fallback = node.getAttribute('ContentType') ?? undefined;
	}
	return fallback;
}

export interface VisioOlePreview {
	object: VisioOleObject;
	image?: VisioImage;
	vector?: VisioForeignVector;
}

/**
 * An embedded OLE object (Foreign shape, ForeignType Object) that is not a chart or an equation:
 * Visio draws it from the picture it saved with the object (an `image` relationship of the
 * embedded part, EMF or a bitmap), as an icon or as its content as stored. The object itself is
 * never activated or parsed here; its part is kept byte for byte by every edit, and an embedded
 * Word, Excel, PowerPoint or Visio file can be read back with `readVisioEmbeddedObject` to open.
 */
export async function readEmbeddedOleObject(
	pkg: VisioPackage,
	sourcePart: string,
	foreignData: Element,
	report: Report,
	options?: VisioImageOptions,
	converter?: VisioMetafileTreeConverter,
): Promise<VisioOlePreview | undefined> {
	const refs = children(foreignData, 'Rel');
	const id = refs.length === 1 ? refs[0]?.getAttributeNS(NS.r, 'id') : undefined;
	const rel = id ? (await pkg.relationships(sourcePart)).get(id) : undefined;
	if (!rel) {
		report('ole-object-relationship', 'The embedded object has no unambiguous relationship.');
		return undefined;
	}
	if (rel.mode !== 'Internal') {
		report('ole-object-linked', 'Linked OLE objects are not loaded; the link is kept as it is.');
		return undefined;
	}
	if ((rel.type !== PACKAGE && rel.type !== OLE_OBJECT) || !pkg.has(rel.target)) {
		report('ole-object-relationship', 'The embedded object does not reference an embedded part.');
		return undefined;
	}
	const format = classifyVisioOleEmbedding(
		rel.target,
		rel.type,
		await contentTypeOf(pkg, rel.target),
	);
	const refusal = REFUSALS[format];
	const object: VisioOleObject = {
		part: rel.target,
		format,
		fileName: rel.target.slice(rel.target.lastIndexOf('/') + 1),
		showAsIcon: foreignData.getAttribute('ShowAsIcon') === '1',
		...(refusal ? { refusal } : {}),
	};
	const preview = [...(await pkg.relationships(rel.target)).values()].find(
		(candidate) => candidate.type === IMAGE && candidate.mode === 'Internal',
	);
	if (!preview || !pkg.has(preview.target)) {
		report(
			'ole-object-no-preview',
			'The embedded object has no saved picture, so it is not drawn; it is kept as it is.',
		);
		return { object };
	}
	const target = preview.target.toLowerCase();
	if (target.endsWith('.emf')) {
		const vector = await inspectVisioMetafilePart(pkg, preview.target, report, converter);
		return vector ? { object, vector } : { object };
	}
	if (target.endsWith('.wmf')) {
		report(
			'ole-object-wmf-preview',
			'The embedded object is saved with a WMF picture, which is not drawn yet; it is kept as it is.',
		);
		return { object };
	}
	const image = await readVisioImagePart(pkg, preview.target, report, options);
	return image ? { object, image } : { object };
}

/** The largest embedded document `readVisioEmbeddedObject` hands back to open. */
export const VISIO_EMBEDDED_OPEN_MAX_BYTES = 32 * 1024 * 1024;
const OPENABLE = new Set<VisioOleObjectFormat>(['docx', 'xlsx', 'pptx', 'vsdx']);

/** True when an embedded object can be opened as a document of its own (Open in its editor). */
export const isOpenableVisioOleObject = (object: VisioOleObject): boolean =>
	OPENABLE.has(object.format) && !object.refusal;

/**
 * The bytes of an embedded Word, Excel, PowerPoint or Visio document in `vsdx`, for opening it in
 * its own editor (Visio's Open on an embedded object). Anything else is refused: OLE containers,
 * macro-enabled documents, parts that are not ZIP packages and parts over 32 MiB.
 */
export async function readVisioEmbeddedObject(
	vsdx: Uint8Array,
	part: string,
): Promise<{ bytes: Uint8Array; fileName: string; format: VisioOleObjectFormat }> {
	const pkg = await VisioPackage.open(vsdx);
	const refuse = (message: string): never => {
		return fail('EDIT_TARGET_NOT_FOUND', message);
	};
	if (!pkg.has(part)) refuse('The embedded object is not in this drawing.');
	const format = classifyVisioOleEmbedding(part, PACKAGE, await contentTypeOf(pkg, part));
	if (!OPENABLE.has(format)) refuse(REFUSALS[format] ?? REFUSALS.other!);
	if (pkg.getPartByteLength(part) > VISIO_EMBEDDED_OPEN_MAX_BYTES)
		refuse('The embedded document is larger than 32 MiB and is not opened here.');
	const bytes = await pkg.readBytes(part);
	if (bytes[0] !== 0x50 || bytes[1] !== 0x4b || bytes[2] !== 0x03 || bytes[3] !== 0x04)
		refuse('The embedded document is not a valid Office package.');
	return { bytes: Uint8Array.from(bytes), fileName: part.slice(part.lastIndexOf('/') + 1), format };
}
