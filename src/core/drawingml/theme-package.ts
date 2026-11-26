import JSZip from 'jszip';
import { buildContentTypesXml, contentTypeForPart, parseContentTypes } from '../opc/content-types';
import { isPartName, partNameKey } from '../opc/part-names';
import {
	buildRelationshipsXml,
	parseRelationshipsPart,
	relationshipsPartFor,
	resolvePartPath,
} from '../opc/relationships';
import { openZip } from '../opc/zip-limits';
import { NS, elements, named, parseXml, type XmlElement } from '../xml/index';
import { resolveDrawingColor } from './drawing-color';
import { parseTheme } from './theme';
import { THEME_COLOR_SLOTS } from './theme-model';

export const MAX_THEME_PACKAGE_BYTES = 8 * 1024 * 1024;
const THEME_CT = 'application/vnd.openxmlformats-officedocument.theme+xml';
const MANAGER_CT = 'application/vnd.openxmlformats-officedocument.themeManager+xml';
const OFFICE_DOCUMENT_REL = `${NS.r}/officeDocument`;
const THEME_REL = `${NS.r}/theme`;
const MANAGER = 'theme/theme/themeManager.xml';
const THEME = 'theme/theme/theme1.xml';
const XML = {
	label: 'Office theme',
	warningsAsErrors: true,
	limits: { maxDepth: 40, maxNodes: 5000 },
};

function expectChildren(parent: XmlElement, names: readonly string[]) {
	const items = elements(parent);
	for (const [index, name] of names.entries()) {
		if (!named(items[index], name, NS.a, false)) throw new Error(`Office theme: expected ${name}.`);
	}
	return items;
}

function validateReferences(root: XmlElement): void {
	for (const attribute of Array.from(root.attributes)) {
		if (attribute.namespaceURI === NS.r)
			throw new Error('Office theme: related resources are unsupported.');
	}
	for (const child of elements(root)) validateReferences(child);
}

function validateTheme(xml: string) {
	if (xml.length > 512 * 1024) throw new Error('Office theme XML is too large.');
	const root = parseXml(xml, XML).documentElement;
	if (!named(root, 'theme', NS.a, false)) throw new Error('Office theme: expected theme.');
	const [holder] = expectChildren(root, ['themeElements']);
	if (!holder) throw new Error('Office theme: missing theme elements.');
	const [colors, fonts, format] = expectChildren(holder, ['clrScheme', 'fontScheme', 'fmtScheme']);
	if (!colors || !fonts || !format) throw new Error('Office theme: missing schemes.');
	for (const slot of expectChildren(colors, THEME_COLOR_SLOTS).slice(0, 12)) {
		const choice = elements(slot);
		if (choice.length !== 1 || choice[0]?.namespaceURI !== NS.a)
			throw new Error('Office theme: invalid colour slot.');
	}
	for (const collection of expectChildren(fonts, ['majorFont', 'minorFont']).slice(0, 2)) {
		for (const face of expectChildren(collection, ['latin', 'ea', 'cs'])) {
			if (face.namespaceURI === NS.a && (face.getAttribute('typeface')?.length ?? 0) > 100)
				throw new Error('Office theme: font name is too long.');
		}
	}
	const lists = expectChildren(format, [
		'fillStyleLst',
		'lnStyleLst',
		'effectStyleLst',
		'bgFillStyleLst',
	]);
	for (const list of lists.slice(0, 4)) {
		const entries = elements(list);
		if (
			entries.length < 3 ||
			entries.length > 32 ||
			entries.some((item) => item.namespaceURI !== NS.a)
		)
			throw new Error('Office theme: invalid style matrix.');
	}
	for (const effect of elements(lists[2]!)) {
		const [entry] = elements(effect);
		if (
			!named(effect, 'effectStyle', NS.a, false) ||
			(!named(entry, 'effectLst', NS.a, false) && !named(entry, 'effectDag', NS.a, false))
		)
			throw new Error('Office theme: invalid effect style.');
	}
	validateReferences(root);
	const theme = parseTheme(root);
	if (
		THEME_COLOR_SLOTS.some((slot) => {
			const color = theme.colorScheme.colors[slot];
			return !color || !resolveDrawingColor(color);
		}) ||
		!theme.formatScheme ||
		theme.formatScheme.lineStyles.length < 3 ||
		theme.formatScheme.effectStyles.length < 3
	)
		throw new Error('Office theme: incomplete schemes.');
	return theme;
}

function relationshipTarget(xml: string, source: string, type: string): string {
	parseXml(xml, XML);
	const matches = parseRelationshipsPart(xml).filter((relationship) => relationship.type === type);
	const [relationship] = matches;
	if (matches.length !== 1 || relationship?.mode !== 'Internal')
		throw new Error('Office theme: expected a unique internal relationship.');
	const target = relationship.target;
	// The two required relationships point to parts at or below their source directory.
	if (!isPartName(target) || target.includes(':'))
		throw new Error('Office theme: invalid part target.');
	return resolvePartPath(source, target);
}

/** Reads the default DrawingML theme from a .thmx package, without loading its ancillary parts. */
export async function readThemePackage(bytes: Uint8Array | ArrayBuffer) {
	const zip = await openZip(bytes, {
		label: 'Office theme',
		maxInputBytes: MAX_THEME_PACKAGE_BYTES,
		maxEntryBytes: 8 * 1024 * 1024,
		maxTotalBytes: 24 * 1024 * 1024,
		maxEntries: 256,
	});
	const entries = new Map<string, string>();
	for (const { name, dir } of Object.values(zip.files)) {
		if (dir) continue;
		const key = partNameKey(name);
		if (entries.has(key)) throw new Error('Office theme: duplicate part name.');
		entries.set(key, name);
	}
	const read = async (name: string) => {
		const item = entries.get(partNameKey(name));
		if (item === undefined) throw new Error('Office theme: missing required part.');
		const xml = await zip.file(item)?.async('string');
		if (!xml || xml.length > 512 * 1024) throw new Error('Office theme: invalid XML part size.');
		return xml;
	};
	const manifest = await read('[Content_Types].xml');
	const root = parseXml(manifest, XML).documentElement;
	if (!named(root, 'Types', NS.contentTypes, false))
		throw new Error('Office theme: invalid content types.');
	const types = parseContentTypes(manifest);
	if (
		elements(root).some(
			(node) =>
				node.namespaceURI !== NS.contentTypes ||
				(node.localName !== 'Default' && node.localName !== 'Override'),
		) ||
		elements(root).length !== types.defaults.size + types.overrides.size ||
		types.overrides.size !== new Set([...types.overrides.keys()].map(partNameKey)).size
	)
		throw new Error('Office theme: invalid content type declarations.');
	const manager = relationshipTarget(await read('_rels/.rels'), '', OFFICE_DOCUMENT_REL);
	if (
		contentTypeForPart(types, manager) !== MANAGER_CT ||
		!named(parseXml(await read(manager), XML).documentElement, 'themeManager', NS.a, false)
	)
		throw new Error('Office theme: invalid theme manager.');
	const part = relationshipTarget(await read(relationshipsPartFor(manager)), manager, THEME_REL);
	if (contentTypeForPart(types, part) !== THEME_CT)
		throw new Error('Office theme: invalid theme type.');
	return validateTheme(await read(part));
}

/** Packages a self-contained DrawingML theme part in the standalone Office theme envelope. */
export async function writeThemePackage(themeXml: string): Promise<Uint8Array> {
	validateTheme(themeXml);
	const zip = new JSZip();
	const rels = (type: string, target: string) =>
		buildRelationshipsXml(new Map([['rId1', { type, target, mode: 'Internal' }]]));
	zip.file(
		'[Content_Types].xml',
		buildContentTypesXml({
			defaults: new Map([['rels', 'application/vnd.openxmlformats-package.relationships+xml']]),
			overrides: new Map([
				[`/${MANAGER}`, MANAGER_CT],
				[`/${THEME}`, THEME_CT],
			]),
		}),
	);
	zip.file('_rels/.rels', rels(OFFICE_DOCUMENT_REL, MANAGER));
	zip.file(MANAGER, `<a:themeManager xmlns:a="${NS.a}"/>`);
	zip.file(relationshipsPartFor(MANAGER), rels(THEME_REL, 'theme1.xml'));
	zip.file(THEME, themeXml);
	return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
