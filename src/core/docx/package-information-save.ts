// Writes the package metadata of `PackageInformation` back on save: a field the model no longer
// has is removed from its part, as Word's Inspect Document removes it.
import type JSZip from 'jszip';
import {
	parseAppProperties,
	parseCoreProperties,
	parseCustomProperties,
	parseRelationships,
	relationshipsPartFor,
	removeContentTypeOverride,
	resolvePartPath,
	writeAppProperties,
	writeCoreProperties,
	writeCustomProperties,
} from '../opc/index';
import { buildXml, parseXml, type XmlElement } from './xml';
import { applySettingsFlag } from './settings';
import { W15_NS } from './comment-parts';
import {
	ATTACHED_TEMPLATE_TYPE,
	DOCUMENT_RELS,
	PEOPLE_TYPE,
	SETTINGS_PART,
	SETTINGS_RELS,
	settingsChild,
	type PackageInformation,
} from './package-information';

const CORE = 'docProps/core.xml';
const APP = 'docProps/app.xml';
const CUSTOM = 'docProps/custom.xml';
const ROOT_RELS = '_rels/.rels';

/** Removes the `<Relationship/>` tags `drop` selects from a `.rels` part, keeping the rest as written. */
export async function removeRelationships(
	zip: JSZip,
	relsPath: string,
	drop: (type: string, target: string) => boolean,
): Promise<void> {
	const xml = await zip.file(relsPath)?.async('string');
	if (!xml) return;
	const rels = parseRelationships(xml);
	let next = xml;
	for (const tag of xml.match(/<Relationship\b[^>]*\/>/g) ?? []) {
		const id = /\bId="([^"]*)"/.exec(tag)?.[1];
		const rel = id ? rels.get(id) : undefined;
		if (rel && drop(rel.type, rel.target)) next = next.replace(tag, '');
	}
	if (next === xml) return;
	if (/<Relationship\b/.test(next)) zip.file(relsPath, next);
	else zip.remove(relsPath);
}

/** Deletes a part with its relationships part and content-type override. */
export async function removePart(zip: JSZip, part: string): Promise<void> {
	zip.remove(part);
	zip.remove(relationshipsPartFor(part));
	await removeContentTypeOverride(zip, part);
}

const changed = (a: unknown, b: unknown) => JSON.stringify(a ?? null) !== JSON.stringify(b ?? null);

async function patchProperties(zip: JSZip, next: PackageInformation, prior: PackageInformation) {
	const core = await zip.file(CORE)?.async('string');
	if (
		core &&
		(changed(next.category, prior.category) || changed(next.contentStatus, prior.contentStatus))
	) {
		const props = parseCoreProperties(core);
		for (const key of ['category', 'contentStatus'] as const) {
			const value = next[key];
			if (value) props[key] = value;
			else delete props[key];
		}
		zip.file(CORE, writeCoreProperties(props, core));
	}
	const appKeys = ['company', 'manager', 'hyperlinkBase', 'template'] as const;
	const app = await zip.file(APP)?.async('string');
	if (app && appKeys.some((key) => changed(next[key], prior[key]))) {
		const props = parseAppProperties(app);
		for (const key of appKeys) {
			const value = next[key];
			if (value) props[key] = value;
			else delete props[key];
		}
		zip.file(APP, writeAppProperties(props, app));
	}
	const custom = await zip.file(CUSTOM)?.async('string');
	if (custom && changed(next.customProperties, prior.customProperties)) {
		const keep = new Set(next.customProperties ?? []);
		const xml = writeCustomProperties(
			parseCustomProperties(custom).filter((property) => keep.has(property.name)),
		);
		if (xml) zip.file(CUSTOM, xml);
		else {
			await removePart(zip, CUSTOM);
			await removeRelationships(zip, ROOT_RELS, (_type, target) =>
				[CUSTOM, `/${CUSTOM}`].includes(target),
			);
		}
	}
}

/** Custom XML data parts the model dropped: each item, its properties part and relationships. */
async function removeCustomXml(zip: JSZip, next: PackageInformation, prior: PackageInformation) {
	const keep = new Set(next.customXmlParts ?? []);
	const dropped = (prior.customXmlParts ?? []).filter((part) => !keep.has(part));
	if (!dropped.length) return;
	for (const part of dropped) {
		const rels = parseRelationships(await zip.file(relationshipsPartFor(part))?.async('string'));
		for (const rel of rels.values())
			if (rel.mode === 'Internal') await removePart(zip, resolvePartPath(part, rel.target));
		await removePart(zip, part);
	}
	await removeRelationships(zip, DOCUMENT_RELS, (_type, target) =>
		dropped.includes(resolvePartPath('word/document.xml', target)),
	);
}

/** settings.xml: the attached template and `w15:docId`, removed when the model dropped them. */
async function patchSettings(zip: JSZip, next: PackageInformation, prior: PackageInformation) {
	const dropTemplate = Boolean(prior.attachedTemplate && !next.attachedTemplate);
	const dropDocId = Boolean(prior.documentId && !next.documentId);
	const xml = await zip.file(SETTINGS_PART)?.async('string');
	if (xml && (dropTemplate || dropDocId)) {
		const doc = parseXml(xml);
		const root = doc.documentElement as XmlElement;
		const remove = (element: XmlElement | undefined) => element && root.removeChild(element);
		if (dropTemplate) remove(settingsChild(root, 'attachedTemplate'));
		if (dropDocId) remove(settingsChild(root, 'docId', W15_NS));
		zip.file(SETTINGS_PART, buildXml(doc));
	}
	if (dropTemplate)
		await removeRelationships(zip, SETTINGS_RELS, (type) => type === ATTACHED_TEMPLATE_TYPE);
	for (const key of ['removePersonalInformation', 'removeDateAndTime'] as const)
		if (Boolean(next[key]) !== Boolean(prior[key]))
			await applySettingsFlag(zip, key, Boolean(next[key]));
}

/** Applies the differences between the loaded package metadata and the model's. */
export async function applyPackageInformation(
	zip: JSZip,
	next: PackageInformation | undefined,
	prior: PackageInformation | undefined,
): Promise<void> {
	const after = next ?? {};
	const before = prior ?? {};
	if (!changed(after, before)) return;
	await patchProperties(zip, after, before);
	await removeCustomXml(zip, after, before);
	await patchSettings(zip, after, before);
	if (before.people && !after.people) {
		await removePart(zip, 'word/people.xml');
		await removeRelationships(zip, DOCUMENT_RELS, (type) => type === PEOPLE_TYPE);
	}
}
