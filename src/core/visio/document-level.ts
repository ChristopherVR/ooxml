// The part of a read that depends on the document, not on a page: the colours, fonts, themes and
// styles, the print-setting reader and the masters with their shapes. A caller that reads the same
// package again after only page parts changed passes a cache (see {@link VisioDocumentCache}), and
// this part is reused when every part it read is byte-for-byte the one it read before.
import type { VisioMaster } from './model';
import { type MasterRecord } from './master-list';
import { metadata, metadataAttributes } from './metadata';
import { VisioPackageError, type VisioPackage } from './package';
import { indexedPart, related, visioXml } from './parts';
import { prepareImages } from './prepare-images';
import { createVisioPrintSettingsReader } from './print-settings';
import {
	attribute,
	child,
	children,
	readShapes,
	readSheet,
	type RawShape,
	type Report,
} from './sheet';
import type { Resources } from './style';
import type { StyleRecord } from './style-inheritance';
import { loadVisioThemes } from './theme';
import type { ParseVsdxOptions } from './parser';

export interface DocumentLevel {
	documentRoot: Element;
	resources: Resources;
	styles: Map<string, StyleRecord>;
	readPrintSettings: ReturnType<typeof createVisioPrintSettingsReader>;
	masters: Map<string, RawShape[]>;
	masterRecords: MasterRecord[];
	/** The diagnostics the read reported, in order, so a reuse reports them again. */
	reports: Parameters<Report>[];
}

/** What the read consumed from its limits once the masters were listed. */
export interface ListedMasters {
	documentMasters: VisioMaster[];
	counters: {
		curveWork: number;
		geometry: number;
		text: number;
		expansion: number;
		textRuns: number;
		paragraphs: number;
		shapes: number;
		metadata: [number, number];
		layers: [number, number];
	};
}

interface CachedLevel {
	/** The options that shaped the read. */
	options: string;
	converter: unknown;
	/** Each part the read looked at with a print of it (declared CRC-32 and size). */
	reads: [string, string][];
	level: DocumentLevel;
	listed: ListedMasters;
}

/**
 * Keeps the document-level data of the last read. Hand the same object to every read of one
 * drawing: a read whose package has identical parts where the last one looked (and the same
 * options) reuses the data instead of reading styles, themes and masters again; anything else
 * reads them afresh and replaces the entry. Identity of parts is judged by the CRC-32 and size
 * their ZIP entries declare, which the package validates when a part is inflated.
 */
export interface VisioDocumentCache {
	entry?: CachedLevel;
}

export function optionsKey(options: ParseVsdxOptions): string {
	const { metafileConverter: _converter, pages: _pages, documentCache: _cache, ...rest } = options;
	return JSON.stringify(rest);
}

export const cacheMatches = (
	entry: CachedLevel | undefined,
	pkg: VisioPackage,
	options: ParseVsdxOptions,
): entry is CachedLevel =>
	entry !== undefined &&
	entry.options === optionsKey(options) &&
	entry.converter === options.metafileConverter &&
	entry.reads.every(([path, print]) => pkg.printOf(path) === print);

export type { CachedLevel };

export async function readDocumentLevel(
	pkg: VisioPackage,
	documentPart: string,
	options: ParseVsdxOptions,
	checkTime: () => void,
	forward: Report,
): Promise<DocumentLevel> {
	const reports: Parameters<Report>[] = [];
	const report: Report = (...args) => {
		reports.push(args);
		forward(...args);
	};
	const documentRoot = await visioXml(pkg, documentPart, 'VisioDocument');
	const resources: Resources = {
		colors: new Map<string, string>(),
		fonts: new Map<string, string>(),
		themes: await loadVisioThemes(pkg, documentPart, report),
	};
	for (const color of children(child(documentRoot, 'Colors'), 'ColorEntry'))
		resources.colors.set(attribute(color, 'IX') ?? '', attribute(color, 'RGB') ?? '');
	const faceNames = children(child(documentRoot, 'FaceNames'), 'FaceName');
	if (faceNames.length > 10000)
		throw new VisioPackageError(
			'METADATA_LIMIT',
			'Document font count exceeds its metadata limit.',
		);
	let fontCharacters = 0;
	faceNames.forEach((font, index) => {
		checkTime();
		fontCharacters += (attribute(font, 'Name') ?? attribute(font, 'NameU') ?? 'Arial').length;
		if (fontCharacters > 1_000_000)
			throw new VisioPackageError(
				'METADATA_LIMIT',
				'Document fonts exceed aggregate metadata limits.',
			);
		resources.fonts.set(
			metadata(attribute(font, 'ID') ?? String(index), 256, 'Font ID'),
			metadata(attribute(font, 'Name') ?? attribute(font, 'NameU') ?? 'Arial', 1024, 'Font name'),
		);
	});
	const styles = new Map<string, StyleRecord>();
	for (const style of children(child(documentRoot, 'StyleSheets'), 'StyleSheet'))
		styles.set(metadata(attribute(style, 'ID') ?? '', 256, 'Style ID'), {
			sheet: readSheet(style),
			attributes: metadataAttributes(style),
		});
	const rootStyle = styles.get('0');
	if (rootStyle?.attributes.get('NameU') === 'No Style') resources.rootSheet = rootStyle.sheet;
	const printStyleGroups = children(documentRoot, 'StyleSheets');
	const readPrintSettings = createVisioPrintSettingsReader(
		children(documentRoot, 'DocumentSheet'),
		printStyleGroups.length === 1
			? children(printStyleGroups[0], 'StyleSheet').filter((node) => attribute(node, 'ID') === '0')
			: [],
	);
	const masters = new Map<string, RawShape[]>();
	const masterRecords: MasterRecord[] = [];
	const mastersPart = await related(pkg, documentPart, 'masters', false);
	if (mastersPart) {
		for (const master of children(await visioXml(pkg, mastersPart, 'Masters'), 'Master')) {
			const part = await indexedPart(pkg, mastersPart, master, 'master');
			const id = metadata(attribute(master, 'ID') ?? '', 256, 'Master ID');
			if (!id || masters.has(id))
				throw new VisioPackageError('INVALID_MASTER_ID', 'Master IDs must be present and unique.');
			const masterShapes = readShapes(await visioXml(pkg, part, 'MasterContents'));
			await prepareImages(
				masterShapes,
				pkg,
				part,
				report,
				options.images,
				options.metafileConverter,
			);
			masters.set(id, masterShapes);
			masterRecords.push({ id, node: master, shapes: masterShapes });
		}
	}
	return { documentRoot, resources, styles, readPrintSettings, masters, masterRecords, reports };
}
