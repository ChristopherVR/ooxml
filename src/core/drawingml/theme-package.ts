import JSZip from 'jszip';
import { buildContentTypesXml, contentTypeForPart, parseContentTypes } from '../opc/content-types';
import { isPartName, partNameKey } from '../opc/part-names';
import {
	buildRelationshipsXml,
	parseRelationshipsPart,
	relationshipsPartFor,
	resolvePartPath,
} from '../opc/relationships';
import { declaredEntrySize, inflateEntryGuarded, openZip } from '../opc/zip-limits';
import {
	NS,
	buildXml,
	decodeXmlBytes,
	elements,
	named,
	parseXml,
	type XmlElement,
} from '../xml/index';
import { resolveDrawingColor } from './drawing-color';
import { isDrawingFillElement } from './drawing-fill';
import { parseTheme } from './theme';
import { THEME_COLOR_SLOTS } from './theme-model';

export const MAX_THEME_PACKAGE_BYTES = 8 * 1024 * 1024;
const MAX_THEME_XML_BYTES = 512 * 1024;
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

type ThemePackageWarning = 'pictureFillsReplaced';
interface ThemePackageOptions {
	onWarning?: (warning: ThemePackageWarning) => void;
}

function expectChildren(parent: XmlElement, names: readonly string[]) {
	const items = elements(parent);
	for (const [index, name] of names.entries()) {
		if (!named(items[index], name, NS.a, false)) throw new Error(`Office theme: expected ${name}.`);
	}
	return items;
}

function hasReferences(root: XmlElement): boolean {
	return (
		Array.from(root.attributes).some((attribute) => attribute.namespaceURI === NS.r) ||
		elements(root).some(hasReferences)
	);
}

function replacePictureFills(lists: readonly XmlElement[]): boolean {
	let replaced = false;
	for (const list of lists) {
		for (const fill of elements(list)) {
			if (!named(fill, 'blipFill', NS.a, false) || !hasReferences(fill)) continue;
			const doc = list.ownerDocument;
			const solid = doc.createElementNS(NS.a, 'a:solidFill');
			const color = doc.createElementNS(NS.a, 'a:schemeClr');
			color.setAttribute('val', 'phClr');
			solid.appendChild(color);
			list.replaceChild(solid, fill);
			replaced = true;
		}
	}
	return replaced;
}

function validateTheme(xml: string) {
	const doc = parseXml(xml, XML);
	const root = doc.documentElement;
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
			entries.some((item) => {
				if (item.namespaceURI !== NS.a) return true;
				if (list.localName === 'lnStyleLst') return item.localName !== 'ln';
				if (list.localName === 'effectStyleLst') return item.localName !== 'effectStyle';
				return !isDrawingFillElement(item);
			})
		)
			throw new Error('Office theme: invalid style matrix.');
	}
	for (const effect of elements(lists[2]!)) {
		const [entry] = elements(effect);
		if (!named(entry, 'effectLst', NS.a, false) && !named(entry, 'effectDag', NS.a, false))
			throw new Error('Office theme: invalid effect style.');
	}
	const replaced = replacePictureFills(lists.filter((_, index) => index === 0 || index === 3));
	if (hasReferences(root)) throw new Error('Office theme: related resources are unsupported.');
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
	return { theme, xml: replaced ? buildXml(doc) : xml, replaced };
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
export async function readThemePackage(
	bytes: Uint8Array | ArrayBuffer,
	options?: ThemePackageOptions,
) {
	const zip = await openZip(bytes, {
		label: 'Office theme',
		maxInputBytes: MAX_THEME_PACKAGE_BYTES,
		maxEntryBytes: 8 * 1024 * 1024,
		maxTotalBytes: 24 * 1024 * 1024,
		// Office saves the default theme alongside variants with their own slide layouts and media.
		maxEntries: 640,
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
		const entry = item === undefined ? null : zip.file(item);
		if (!entry) throw new Error('Office theme: missing required part.');
		const size = declaredEntrySize(entry);
		if (size === 0 || size > MAX_THEME_XML_BYTES)
			throw new Error('Office theme: invalid XML part size.');
		return decodeXmlBytes(
			await inflateEntryGuarded(entry, 0, { label: XML.label, maxEntryBytes: MAX_THEME_XML_BYTES }),
		);
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
	const { theme, replaced } = validateTheme(await read(part));
	if (replaced) options?.onWarning?.('pictureFillsReplaced');
	return theme;
}

/** Packages a self-contained DrawingML theme part in the standalone Office theme envelope. */
export async function writeThemePackage(
	themeXml: string,
	options?: ThemePackageOptions,
): Promise<Uint8Array> {
	if (themeXml.length > MAX_THEME_XML_BYTES) throw new Error('Office theme XML is too large.');
	// The byte encoding is ours to choose, regardless of the source XML declaration.
	const xml = themeXml.replace(
		/^\uFEFF?<\?xml\s[^?]*\?>/u,
		'<?xml version="1.0" encoding="UTF-8"?>',
	);
	const encoder = new TextEncoder();
	if (encoder.encode(xml).length > MAX_THEME_XML_BYTES)
		throw new Error('Office theme XML is too large.');
	const validated = validateTheme(xml);
	const bytes = encoder.encode(validated.xml);
	if (bytes.length > MAX_THEME_XML_BYTES) throw new Error('Office theme XML is too large.');
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
	zip.file(THEME, bytes);
	const packageBytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
	if (validated.replaced) options?.onWarning?.('pictureFillsReplaced');
	return packageBytes;
}
