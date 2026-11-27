/** Text edits to the package's bookkeeping parts: content types and relationships. */
import { xmlEscape } from '../../lib/package';
import { SAMPLE_AUTHOR, SAMPLE_DATE_ISO } from '../../lib/types';
import { XML_HEADER } from './xml';

export function insertBefore(text: string, marker: string, addition: string): string {
	const at = text.lastIndexOf(marker);
	if (at < 0) throw new Error(`Expected ${marker} in the part.`);
	return text.slice(0, at) + addition + text.slice(at);
}

/** Adds Override entries for parts the content types do not list yet. */
export function withContentTypes(
	types: string,
	entries: readonly (readonly [string, string])[],
): string {
	const missing = entries
		.filter(([part]) => !types.includes(`PartName="${part}"`))
		.map(([part, type]) => `<Override PartName="${part}" ContentType="${type}"/>`);
	return insertBefore(types, '</Types>', missing.join(''));
}

export function withRelationship(rels: string, id: string, type: string, target: string): string {
	if (rels.includes(`Target="${target}"`)) return rels;
	return insertBefore(
		rels,
		'</Relationships>',
		`<Relationship Id="${id}" Type="${type}" Target="${target}"/>`,
	);
}

export const EMPTY_RELS = `${XML_HEADER}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;

const SCHEMAS = 'http://schemas.openxmlformats.org';

/** `docProps/core.xml`: the document properties, with the samples' fixed author and date. */
export function corePropertiesXml(properties: {
	title: string;
	subject: string;
	description: string;
	keywords: readonly string[];
}): string {
	const text = (tag: string, value: string) => `<${tag}>${xmlEscape(value)}</${tag}>`;
	const date = (tag: string) =>
		`<dcterms:${tag} xsi:type="dcterms:W3CDTF">${SAMPLE_DATE_ISO}</dcterms:${tag}>`;
	return (
		`${XML_HEADER}<cp:coreProperties xmlns:cp="${SCHEMAS}/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
		text('dc:title', properties.title) +
		text('dc:subject', properties.subject) +
		text('dc:creator', SAMPLE_AUTHOR) +
		text('cp:keywords', properties.keywords.join(', ')) +
		text('dc:description', properties.description) +
		text('cp:lastModifiedBy', SAMPLE_AUTHOR) +
		date('created') +
		date('modified') +
		'<dc:language>en-US</dc:language></cp:coreProperties>'
	);
}

/** `docProps/app.xml`. */
export const APP_PROPERTIES_XML =
	`${XML_HEADER}<Properties xmlns="${SCHEMAS}/officeDocument/2006/extended-properties">` +
	'<Application>ooxml-core</Application><Company>OOXML samples</Company></Properties>';

export const PROPERTY_CONTENT_TYPES: readonly (readonly [string, string])[] = [
	['/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml'],
	['/docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml'],
];

/** The package relationships with the two property parts added. */
export function withPropertyRelationships(rels: string): string {
	const core = withRelationship(
		rels,
		'rIdCore',
		`${SCHEMAS}/package/2006/relationships/metadata/core-properties`,
		'docProps/core.xml',
	);
	return withRelationship(
		core,
		'rIdApp',
		`${SCHEMAS}/officeDocument/2006/relationships/extended-properties`,
		'docProps/app.xml',
	);
}
