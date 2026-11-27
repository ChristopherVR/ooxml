/**
 * The pass after `saveDocx`: swaps the hand-written fragments in for their sentinels, adds the
 * parts the core does not write for a new document (styles, numbering, settings, extended and
 * custom properties) and the core properties with the sample date. The save itself is
 * reproducible (`save.ts` passes the sample date and a seeded id source).
 */
import {
	writeAppProperties,
	writeCoreProperties,
	writeCustomProperties,
} from '../../../../src/core/opc/properties/index';
import type {
	AppProperties,
	CoreProperties,
	CustomProperty,
} from '../../../../src/core/opc/properties/types';
import { buildPackage, readPackage, type PackageParts } from '../../lib/package';
import { escapeRegExp } from '../../lib/regexp';
import { SAMPLE_AUTHOR, SAMPLE_DATE_ISO } from '../../lib/types';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const WML = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

/** Namespaces the hand-written fragments use; declared once on each part root that needs them. */
const NAMESPACES: Record<string, string> = {
	r: REL,
	m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
	mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
	wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
	a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
	pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
	wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
	w14: 'http://schemas.microsoft.com/office/word/2010/wordml',
	w15: 'http://schemas.microsoft.com/office/word/2012/wordml',
};
/** Prefixes a reader that does not know them may skip (`mc:Ignorable`). */
const IGNORABLE = ['w14', 'w15'];

export interface FinishSpec {
	raw: ReadonlyMap<string, string>;
	styles: string;
	numbering?: string;
	settings: string;
	core: CoreProperties;
	app: AppProperties;
	custom?: CustomProperty[];
	/** A last edit of single parts, by part name. */
	patch?: Record<string, (xml: string) => string>;
}

const decoder = new TextDecoder();

/** Replaces the paragraph that holds `key` with `xml`. */
function replaceBlock(part: string, key: string, xml: string): string {
	const at = part.indexOf(key);
	if (at < 0) return part;
	const opening = /<w:p[ >]/gu;
	let start = -1;
	for (let match = opening.exec(part); match && match.index < at; match = opening.exec(part))
		start = match.index;
	const end = part.indexOf('</w:p>', at) + '</w:p>'.length;
	if (start < 0 || end < at) throw new Error(`No paragraph around ${key}`);
	return part.slice(0, start) + xml + part.slice(end);
}

/** Replaces the run that holds `key` with `xml`. */
function replaceRun(part: string, key: string, xml: string): string {
	const escaped = escapeRegExp(key);
	const pattern = new RegExp(
		`<w:r>(?:<w:rPr>(?:(?!</w:rPr>).)*</w:rPr>)?<w:t(?: xml:space="preserve")?>${escaped}</w:t></w:r>`,
		'u',
	);
	return part.replace(pattern, () => xml);
}

/** Declares on the root every known prefix the part uses and the root does not declare. */
function declareNamespaces(xml: string): string {
	const rootStart = xml.search(/<w:(?:document|hdr|ftr|footnotes|endnotes)\b/u);
	if (rootStart < 0) return xml;
	const rootEnd = xml.indexOf('>', rootStart);
	let root = xml.slice(rootStart, rootEnd);
	const body = xml.slice(rootEnd);
	for (const [prefix, uri] of Object.entries(NAMESPACES))
		if (body.includes(`<${prefix}:`) && !root.includes(`xmlns:${prefix}=`))
			root += ` xmlns:${prefix}="${uri}"`;
	const ignorable = IGNORABLE.filter((prefix) => root.includes(`xmlns:${prefix}=`));
	if (ignorable.length) {
		if (!root.includes('xmlns:mc=')) root += ` xmlns:mc="${NAMESPACES.mc}"`;
		root = root.replace(/ mc:Ignorable="[^"]*"/u, '');
		root += ` mc:Ignorable="${ignorable.join(' ')}"`;
	}
	return xml.slice(0, rootStart) + root + body;
}

const withDeclaration = (xml: string) =>
	xml.startsWith('<?xml') ? xml : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${xml}`;

export async function finish(bytes: Uint8Array, spec: FinishSpec): Promise<Uint8Array> {
	const parts: PackageParts = {};
	for (const [name, data] of await readPackage(bytes))
		parts[name] = /\.(?:xml|rels)$/u.test(name) ? decoder.decode(data) : data;
	const text = (name: string): string => {
		const part = parts[name];
		if (typeof part !== 'string') throw new Error(`The package has no ${name}`);
		return part;
	};

	// Hand-written fragments, in every story the core wrote.
	for (const name of Object.keys(parts)) {
		if (!/^word\/[^/]+\.xml$/u.test(name)) continue;
		let xml = text(name);
		// A fragment can hold sentinels of its own (a text box's paragraphs), so pass until none is left.
		for (let pass = 0; pass < 4 && xml.includes('[[RAW'); pass++)
			for (const [key, fragment] of spec.raw) {
				if (!xml.includes(key)) continue;
				xml = key.startsWith('[[RAWBLOCK:')
					? replaceBlock(xml, key, fragment)
					: replaceRun(xml, key, fragment);
				if (xml.includes(key)) throw new Error(`${name}: ${key} was not replaced`);
			}
		parts[name] = withDeclaration(declareNamespaces(xml));
	}
	const left = Object.entries(parts).find(
		([, data]) => typeof data === 'string' && data.includes('[[RAW'),
	);
	if (left) throw new Error(`${left[0]} still holds a sentinel`);

	const override = (name: string, type: string) => {
		const types = text('[Content_Types].xml');
		if (types.includes(`PartName="/${name}"`)) return;
		parts['[Content_Types].xml'] = types.replace(
			'</Types>',
			`<Override PartName="/${name}" ContentType="${type}"/></Types>`,
		);
	};
	const relate = (rels: string, type: string, target: string) => {
		const xml =
			typeof parts[rels] === 'string'
				? text(rels)
				: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PACKAGE_REL}"></Relationships>`;
		if (xml.includes(`Type="${type}"`)) {
			parts[rels] = xml;
			return;
		}
		const used = [...xml.matchAll(/Id="rId(\d+)"/gu)].map((match) => Number(match[1]));
		const id = `rId${Math.max(0, ...used) + 1}`;
		parts[rels] = xml.replace(
			'</Relationships>',
			`<Relationship Id="${id}" Type="${type}" Target="${target}"/></Relationships>`,
		);
	};

	const documentRels = 'word/_rels/document.xml.rels';
	parts['word/styles.xml'] = spec.styles;
	override('word/styles.xml', `${WML}.styles+xml`);
	relate(documentRels, `${REL}/styles`, 'styles.xml');
	parts['word/settings.xml'] = spec.settings;
	override('word/settings.xml', `${WML}.settings+xml`);
	relate(documentRels, `${REL}/settings`, 'settings.xml');
	if (spec.numbering) {
		parts['word/numbering.xml'] = spec.numbering;
		override('word/numbering.xml', `${WML}.numbering+xml`);
		relate(documentRels, `${REL}/numbering`, 'numbering.xml');
	}

	// Properties: the core's property writers, with the fixed sample instant.
	parts['docProps/core.xml'] = writeCoreProperties({
		creator: SAMPLE_AUTHOR,
		lastModifiedBy: SAMPLE_AUTHOR,
		created: SAMPLE_DATE_ISO,
		modified: SAMPLE_DATE_ISO,
		revision: '1',
		...spec.core,
	});
	override('docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml');
	relate(
		'_rels/.rels',
		'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
		'docProps/core.xml',
	);
	parts['docProps/app.xml'] = writeAppProperties(spec.app);
	override(
		'docProps/app.xml',
		'application/vnd.openxmlformats-officedocument.extended-properties+xml',
	);
	relate('_rels/.rels', `${REL}/extended-properties`, 'docProps/app.xml');
	const custom = spec.custom && writeCustomProperties(spec.custom);
	if (custom) {
		parts['docProps/custom.xml'] = custom;
		override(
			'docProps/custom.xml',
			'application/vnd.openxmlformats-officedocument.custom-properties+xml',
		);
		relate('_rels/.rels', `${REL}/custom-properties`, 'docProps/custom.xml');
	}

	for (const [name, edit] of Object.entries(spec.patch ?? {})) parts[name] = edit(text(name));
	return buildPackage(parts);
}
