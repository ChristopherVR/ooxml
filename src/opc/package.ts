import type JSZip from 'jszip';
import { RELATIONSHIP_TYPES } from './relationship-types.js';
import { nextRelationshipId } from './relationships.js';

const CONTENT_TYPES_PATH = '[Content_Types].xml';
const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const EMPTY_TYPES = `${XML_DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>`;
const EMPTY_RELS = `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;

function attr(tag: string, name: string): string | undefined {
	return new RegExp(`${name}="([^"]*)"`).exec(tag)?.[1];
}

/**
 * Adds an `<Override>` for `partName` to `[Content_Types].xml` unless one exists. Plain string
 * surgery, so the rest of the part stays byte-identical.
 */
export async function ensureContentTypeOverride(
	zip: JSZip,
	partName: string,
	contentType: string,
): Promise<void> {
	const xml = (await zip.file(CONTENT_TYPES_PATH)?.async('string')) ?? EMPTY_TYPES;
	const target = `/${partName}`;
	for (const tag of xml.match(/<Override\b[^>]*\/>/g) ?? [])
		if (attr(tag, 'PartName') === target) return;
	const insertion = `<Override PartName="${target}" ContentType="${contentType}"/>`;
	zip.file(CONTENT_TYPES_PATH, xml.replace('</Types>', `${insertion}</Types>`));
}

/**
 * Returns the id of the relationship of `type` to `target` in the `.rels` part at `relsPath`,
 * creating the relationship (and the part) when it does not exist. Existing XML is left intact.
 */
export async function ensureRelationship(
	zip: JSZip,
	relsPath: string,
	type: string,
	target: string,
): Promise<string> {
	const xml = (await zip.file(relsPath)?.async('string')) ?? EMPTY_RELS;
	const used = new Set<string>();
	for (const tag of xml.match(/<Relationship\b[^>]*\/>/g) ?? []) {
		const id = attr(tag, 'Id');
		if (id) used.add(id);
		if (id && attr(tag, 'Type') === type && attr(tag, 'Target') === target) return id;
	}
	// Ids are numbered after the highest existing rIdN, not the first gap, so a removed
	// relationship's id is never reused within the same save.
	const highest = Math.max(
		0,
		...[...used].map((id) => Number(id.replace(/^rId/, ''))).filter(Number.isFinite),
	);
	let id = `rId${highest + 1}`;
	if (used.has(id)) id = nextRelationshipId(used);
	const insertion = `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`;
	zip.file(relsPath, xml.replace('</Relationships>', `${insertion}</Relationships>`));
	return id;
}

/** The package-level relationship to the main document part (`_rels/.rels`), if present. */
export async function findOfficeDocumentPart(zip: JSZip): Promise<string | undefined> {
	const xml = await zip.file('_rels/.rels')?.async('string');
	for (const tag of xml?.match(/<Relationship\b[^>]*\/>/g) ?? [])
		if (attr(tag, 'Type') === RELATIONSHIP_TYPES.officeDocument) return attr(tag, 'Target');
	return undefined;
}
