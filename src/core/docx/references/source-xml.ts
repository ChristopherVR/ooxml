import { buildXml, children, first, parseXml, type XmlElement } from '../../xml/index';
import { BIBLIOGRAPHY_NS, BIBLIOGRAPHY_XML_LIMIT } from '../bibliography-part';
import {
	SOURCE_FIELDS,
	SOURCE_LIMIT,
	SOURCE_TEXT_LIMIT,
	type Bibliography,
	type BibliographyAuthor,
	type BibliographySource,
	type BibliographyStyle,
} from './source-model';

const STYLE_NAMES: Record<BibliographyStyle, { file: string; version: string }> = {
	APA: { file: 'APASixthEditionOfficeOnline.xsl', version: '6' },
	MLA: { file: 'MLASeventhEditionOfficeOnline.xsl', version: '7' },
	Chicago: { file: 'CHICAGO.XSL', version: '15' },
};

function rootOf(xml: string) {
	if (xml.length > BIBLIOGRAPHY_XML_LIMIT || !xml.includes(BIBLIOGRAPHY_NS)) return undefined;
	try {
		const root = parseXml(xml, { label: 'Bibliography' }).documentElement;
		if (root?.namespaceURI === BIBLIOGRAPHY_NS && root.localName === 'Sources') return root;
	} catch {
		// Opaque custom XML is retained even when it cannot be read as a source list.
	}
	return undefined;
}

const entries = (element: XmlElement, local: string) => children(element, local, BIBLIOGRAPHY_NS);
const child = (element: XmlElement, local: string) => first(element, local, BIBLIOGRAPHY_NS);
const value = (element: XmlElement, local: string) =>
	child(element, local)?.textContent?.slice(0, SOURCE_TEXT_LIMIT) ?? '';

function authorOf(source: XmlElement): BibliographyAuthor | undefined {
	const authors = child(source, 'Author');
	const role = authors && child(authors, 'Author');
	if (!role) return undefined;
	const company = value(role, 'Corporate');
	if (company) return { kind: 'corporate', name: company };
	const list = child(role, 'NameList');
	if (!list) return undefined;
	return {
		kind: 'people',
		people: entries(list, 'Person')
			.slice(0, 50)
			.map((person) => ({
				first: value(person, 'First'),
				middle: value(person, 'Middle'),
				last: value(person, 'Last'),
			})),
	};
}

export function readBibliographyXml(xml: string): Bibliography | undefined {
	const root = rootOf(xml);
	if (!root) return undefined;
	const sources = entries(root, 'Source');
	if (sources.length > SOURCE_LIMIT) return undefined;
	const selected = (
		root.getAttribute('StyleName') ||
		root.getAttribute('SelectedStyle') ||
		''
	).toLowerCase();
	const style = selected.includes('apa')
		? 'APA'
		: selected.includes('mla')
			? 'MLA'
			: selected.includes('chicago')
				? 'Chicago'
				: undefined;
	return {
		...(style ? { style } : {}),
		sources: sources.flatMap((element): BibliographySource[] => {
			const tag = value(element, 'Tag');
			if (!tag) return [];
			const author = authorOf(element);
			const fields: BibliographySource['fields'] = {};
			for (const key of SOURCE_FIELDS) {
				const text = value(element, key);
				if (text) fields[key] = text;
			}
			return [
				{
					tag,
					type: value(element, 'SourceType'),
					fields,
					...(author ? { author } : {}),
					...(['1', 'true'].includes(value(element, 'Temporary')) ? { placeholder: true } : {}),
				},
			];
		}),
	};
}

function append(parent: XmlElement, key: string, text?: string): XmlElement {
	const element = parent.ownerDocument.createElementNS(BIBLIOGRAPHY_NS, `b:${key}`);
	if (text !== undefined) element.textContent = text;
	parent.appendChild(element);
	return element;
}

function patch(parent: XmlElement, key: string, text: string | undefined) {
	const current = child(parent, key);
	if (text) {
		if (current) current.textContent = text;
		else append(parent, key, text);
	} else if (current) parent.removeChild(current);
}

function patchAuthor(element: XmlElement, author: BibliographyAuthor | undefined) {
	const authors = child(element, 'Author');
	const current = authors && child(authors, 'Author');
	if (authors && current) authors.removeChild(current);
	if (!author) return;
	const role = append(authors ?? append(element, 'Author'), 'Author');
	if (author.kind === 'corporate') append(role, 'Corporate', author.name);
	else {
		const list = append(role, 'NameList');
		for (const name of author.people) {
			const person = append(list, 'Person');
			if (name.last) append(person, 'Last', name.last);
			if (name.first) append(person, 'First', name.first);
			if (name.middle) append(person, 'Middle', name.middle);
		}
	}
}

/** Keep unrecognized source properties, author roles, and style attributes from the input part. */
export function writeBibliographyXml(data: Bibliography, original?: string): string {
	const root =
		(original && rootOf(original)) ||
		rootOf(`<?xml version="1.0" encoding="UTF-8"?><b:Sources xmlns:b="${BIBLIOGRAPHY_NS}"/>`);
	if (!root) throw new Error('Could not create bibliography source list');
	const before = original && readBibliographyXml(original);
	const previous = new Map((before ? before.sources : []).map((source) => [source.tag, source]));
	const existing = new Map(
		entries(root, 'Source').map((element) => [value(element, 'Tag'), element]),
	);
	const wanted = new Set(data.sources.map((source) => source.tag));
	for (const [tag, element] of existing) if (tag && !wanted.has(tag)) root.removeChild(element);
	if (data.style && data.style !== (before && before.style)) {
		const { file, version } = STYLE_NAMES[data.style];
		root.setAttribute('SelectedStyle', '\\' + file);
		root.setAttribute('StyleName', data.style);
		root.setAttribute('Version', version);
	}
	for (const source of data.sources) {
		const prior = previous.get(source.tag);
		if (JSON.stringify(prior) === JSON.stringify(source)) continue;
		const element = existing.get(source.tag) ?? append(root, 'Source');
		patch(element, 'Tag', source.tag);
		patch(element, 'SourceType', source.type);
		if (source.placeholder !== prior?.placeholder)
			patch(element, 'Temporary', source.placeholder ? '1' : undefined);
		if (JSON.stringify(source.author) !== JSON.stringify(prior?.author))
			patchAuthor(element, source.author);
		for (const key of SOURCE_FIELDS)
			if (source.fields[key] !== prior?.fields[key]) patch(element, key, source.fields[key]);
	}
	const xml = buildXml(root.ownerDocument);
	if (xml.length > BIBLIOGRAPHY_XML_LIMIT || data.sources.length > SOURCE_LIMIT)
		throw new Error('Bibliography source list exceeds the document limits');
	return xml;
}
