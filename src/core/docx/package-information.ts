// Package metadata Word's Inspect Document reports and removes, beyond what the rest of the model
// holds: extended (`docProps/app.xml`) and custom (`docProps/custom.xml`) properties, custom XML
// data parts, the attached template, the reviewer list and the privacy settings.
import type JSZip from 'jszip';
import {
	parseAppProperties,
	parseCoreProperties,
	parseCustomProperties,
	parseRelationships,
	resolvePartPath,
} from '../opc/index';
import { children as nsChildren } from '../xml/index';
import { parseXml, WORD_NS, type XmlElement } from './xml';
import { W15_NS } from './comment-parts';
import { readBibliographyPart, type BibliographyPart } from './bibliography-part';

/** Absent fields are unset; read once on load and compared with the loaded value on save. */
export interface PackageInformation {
	/** `cp:category` (`docProps/core.xml`). */
	category?: string;
	/** `cp:contentStatus` (`docProps/core.xml`). */
	contentStatus?: string;
	/** `Company` (`docProps/app.xml`). */
	company?: string;
	/** `Manager` (`docProps/app.xml`). */
	manager?: string;
	/** `HyperlinkBase` (`docProps/app.xml`). */
	hyperlinkBase?: string;
	/** `Template` (`docProps/app.xml`): the attached template's file name. */
	template?: string;
	/** The target of `w:attachedTemplate` in settings.xml: the template's path or URL. */
	attachedTemplate?: string;
	/** Names of the custom properties in `docProps/custom.xml`, in order. */
	customProperties?: string[];
	/** Custom XML data parts (`customXml/item1.xml`...) related to the main document. */
	customXmlParts?: string[];
	bibliography?: BibliographyPart;
	/** `word/people.xml` exists: the identities of the people who reviewed the document. */
	people?: boolean;
	/** settings.xml has `w15:docId`, the document's persistent identifier. */
	documentId?: boolean;
	/** `w:removePersonalInformation` (17.15.1.69): Word strips personal information on save. */
	removePersonalInformation?: boolean;
	/** `w:removeDateAndTime` (17.15.1.70): Word strips revision and comment dates on save. */
	removeDateAndTime?: boolean;
}

export const DOCUMENT_PART = 'word/document.xml';
export const DOCUMENT_RELS = 'word/_rels/document.xml.rels';
export const SETTINGS_PART = 'word/settings.xml';
export const SETTINGS_RELS = 'word/_rels/settings.xml.rels';
export const CUSTOM_XML_TYPE =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml';
export const PEOPLE_TYPE = 'http://schemas.microsoft.com/office/2011/relationships/people';
export const ATTACHED_TEMPLATE_TYPE =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate';

const text = async (zip: JSZip, path: string) => zip.file(path)?.async('string');

/** The settings element of `local` (in `ns`, WordprocessingML by default), if there is one. */
export const settingsChild = (
	root: XmlElement,
	local: string,
	ns: string = WORD_NS,
): XmlElement | undefined => nsChildren(root, local, ns)[0];

async function readSettings(zip: JSZip, info: PackageInformation): Promise<void> {
	const xml = await text(zip, SETTINGS_PART);
	if (!xml) return;
	const root = parseXml(xml).documentElement as XmlElement;
	if (settingsChild(root, 'removePersonalInformation')) info.removePersonalInformation = true;
	if (settingsChild(root, 'removeDateAndTime')) info.removeDateAndTime = true;
	if (settingsChild(root, 'docId', W15_NS)) info.documentId = true;
	if (!settingsChild(root, 'attachedTemplate')) return;
	const rels = parseRelationships(await text(zip, SETTINGS_RELS));
	const template = [...rels.values()].find((rel) => rel.type === ATTACHED_TEMPLATE_TYPE);
	if (template?.target) info.attachedTemplate = template.target;
}

/** Reads the package metadata, or undefined when the package has none of it. */
export async function readPackageInformation(zip: JSZip): Promise<PackageInformation | undefined> {
	const info: PackageInformation = {};
	const core = parseCoreProperties(await text(zip, 'docProps/core.xml'));
	if (core.category) info.category = core.category;
	if (core.contentStatus) info.contentStatus = core.contentStatus;
	const app = parseAppProperties(await text(zip, 'docProps/app.xml'));
	for (const key of ['company', 'manager', 'hyperlinkBase', 'template'] as const) {
		const value = app[key];
		if (value) info[key] = value;
	}
	const custom = parseCustomProperties(await text(zip, 'docProps/custom.xml'));
	if (custom.length) info.customProperties = custom.map((property) => property.name);
	const rels = [...parseRelationships(await text(zip, DOCUMENT_RELS)).values()];
	const customXml = rels
		.filter((rel) => rel.type === CUSTOM_XML_TYPE && rel.mode === 'Internal')
		.map((rel) => resolvePartPath(DOCUMENT_PART, rel.target))
		.filter((part) => zip.file(part));
	if (customXml.length) {
		info.customXmlParts = customXml;
		const bibliography = await readBibliographyPart(zip, customXml);
		if (bibliography) info.bibliography = bibliography;
	}
	if (zip.file('word/people.xml')) info.people = true;
	await readSettings(zip, info);
	return Object.keys(info).length ? info : undefined;
}
