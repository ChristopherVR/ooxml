import JSZip from 'jszip';
import { buildBlankPresentationArchive } from './PresentationBuilder';
import { PptxHandler } from '../../PptxHandler';
import { clonePartGraph, relativePartTarget } from '../../../../opc/clone-part-graph';
import {
	buildRelationshipsXml,
	nextRelationshipId,
	parseRelationships,
	relationshipsPartFor,
	resolvePartPath,
} from '../../../../opc/relationships';
import { buildXml, elements, first, parseXml, type XmlElement } from '../../../../xml/index';

export interface ImportSlidePackageOptions {
	/** Source slide indexes in clipboard order. Defaults to every source slide. */
	slideIndices?: readonly number[];
	/** Destination insertion point, zero based. Defaults to after its last slide. */
	insertAt?: number;
}
export interface ImportSlidePackageResult {
	bytes: Uint8Array;
	/** Imported slide package paths, in the requested order. */
	slidePaths: string[];
}

/**
 * Import slide OOXML and its complete owned dependency graph into another deck.
 * Source layouts, masters and themes receive independent paths. Destination theme,
 * slide size and existing parts remain unchanged. The input archives are immutable.
 * Unsupported global dependencies are rejected explicitly before producing a result.
 */
export async function importSlidePackage(
	destination: ArrayBuffer | Uint8Array,
	source: ArrayBuffer | Uint8Array,
	options: ImportSlidePackageOptions = {},
): Promise<ImportSlidePackageResult> {
	if (
		options.insertAt !== undefined &&
		(!Number.isInteger(options.insertAt) || !Number.isFinite(options.insertAt))
	)
		throw new Error('Insertion point must be a finite integer');
	const [zip, sourceZip] = await Promise.all([
		JSZip.loadAsync(destination),
		JSZip.loadAsync(source),
	]);
	const presentationPath = 'ppt/presentation.xml';
	const sourceDoc = parseXml(await requiredXml(sourceZip, presentationPath));
	const destinationDoc = parseXml(await requiredXml(zip, presentationPath));
	const sourceRoot = sourceDoc.documentElement,
		root = destinationDoc.documentElement;
	const ns = root.namespaceURI ?? '';
	if (sourceRoot.namespaceURI !== ns)
		throw new Error('Cannot import slides between Strict and Transitional packages');
	const sourceSize = first(sourceRoot, 'sldSz', ns),
		size = first(root, 'sldSz', ns);
	if (
		sourceSize?.getAttribute('cx') !== size?.getAttribute('cx') ||
		sourceSize?.getAttribute('cy') !== size?.getAttribute('cy')
	)
		throw new Error('Cannot import slides with a different slide size');
	if (first(sourceRoot, 'embeddedFontLst', ns))
		throw new Error('Cross-deck embedded fonts are not supported');
	const relationshipNs =
		ns === 'http://purl.oclc.org/ooxml/presentationml/main'
			? 'http://purl.oclc.org/ooxml/officeDocument/relationships'
			: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
	const presentationRels = relationshipsPartFor(presentationPath);
	const sourceRels = parseRelationships(await requiredXml(sourceZip, presentationRels));
	for (const relationship of sourceRels.values()) {
		if (!relationship.type.endsWith('/tableStyles') || relationship.mode === 'External') continue;
		const styles = parseXml(
			await requiredXml(sourceZip, resolvePartPath(presentationPath, relationship.target)),
		);
		if (elements(styles.documentElement).some((node) => node.localName === 'tblStyle'))
			throw new Error('Cross-deck custom table styles are not supported');
	}
	const destinationRels = parseRelationships(await requiredXml(zip, presentationRels));
	const masterIds = Array.from(root.getElementsByTagNameNS(ns, 'sldMasterId')).map((entry) =>
		entry.getAttribute('id'),
	);
	for (const relationship of destinationRels.values()) {
		if (!relationship.type.endsWith('/slideMaster') || relationship.mode === 'External') continue;
		const master = parseXml(
			await requiredXml(zip, resolvePartPath(presentationPath, relationship.target)),
		);
		for (const entry of Array.from(master.getElementsByTagNameNS(ns, 'sldLayoutId')))
			masterIds.push(entry.getAttribute('id'));
	}
	const allocateMasterId = nativeIdAllocator(masterIds, 2147483648, 4294967295, 'master/layout');
	const masterParts = new Map<string, 'slideMaster' | 'notesMaster'>();
	for (const relationship of sourceRels.values()) {
		const type = relationship.type.slice(relationship.type.lastIndexOf('/') + 1);
		if ((type === 'slideMaster' || type === 'notesMaster') && relationship.mode !== 'External')
			masterParts.set(resolvePartPath(presentationPath, relationship.target), type);
	}
	const sourceSlides = elements(first(sourceRoot, 'sldIdLst', ns) ?? sourceRoot).filter(
		(node) => node.localName === 'sldId',
	);
	const indexes = options.slideIndices ?? sourceSlides.map((_, index) => index);
	if (
		indexes.some(
			(index) => !Number.isInteger(index) || index < 0 || index >= sourceSlides.length,
		) ||
		new Set(indexes).size !== indexes.length
	)
		throw new Error('Source slide indexes must be unique valid integers');
	if (!indexes.length) return { bytes: new Uint8Array(destination), slidePaths: [] };
	const selected = indexes.map((index) => {
		const rel = sourceRels.get(
			sourceSlides[index]!.getAttributeNS(relationshipNs, 'id') ??
				sourceSlides[index]!.getAttribute('r:id') ??
				'',
		);
		if (!rel || rel.mode === 'External') throw new Error('Source slide relationship is missing');
		return resolvePartPath(presentationPath, rel.target);
	});
	const selectedSet = new Set(selected),
		copies = new Map<string, string>();
	const reserved = new Set(Object.keys(zip.files));
	for (const path of selected) {
		let index = 1;
		while (reserved.has(`ppt/slides/slide${index}.xml`)) index++;
		const newPath = `ppt/slides/slide${index}.xml`;
		reserved.add(newPath);
		copies.set(path, newPath);
		const bytes = await sourceZip.file(path)?.async('uint8array');
		if (!bytes) throw new Error(`Source slide is missing: ${path}`);
		zip.file(newPath, bytes);
	}
	for (const path of selected) {
		await clonePartGraph({
			zip,
			sourceZip,
			copies,
			sourcePart: path,
			targetPart: copies.get(path)!,
			shouldClone: (rel, part) => {
				const type = rel.type.slice(rel.type.lastIndexOf('/') + 1);
				if (type === 'slideMaster' || type === 'notesMaster') masterParts.set(part, type);
				if (/comment/iu.test(type)) throw new Error('Cross-deck comments are not supported');
				if (type === 'slide' && !selectedSet.has(part))
					throw new Error('An imported slide links to a slide outside the selected set');
				return true;
			},
		});
	}
	const slideList = first(root, 'sldIdLst', ns) ?? destinationDoc.createElementNS(ns, 'p:sldIdLst');
	if (!slideList.parentNode) root.insertBefore(slideList, first(root, 'sldSz', ns) ?? null);
	const existingSlides = elements(slideList);
	const insertAt = Math.min(
		Math.max(0, options.insertAt ?? existingSlides.length),
		existingSlides.length,
	);
	const allocateSlideId = nativeIdAllocator(
		existingSlides.map((node) => node.getAttribute('id')),
		256,
		2147483647,
		'slide',
	);
	const used = new Set(destinationRels.keys());
	const register = (part: string, type: string): string => {
		const id = nextRelationshipId(used);
		used.add(id);
		destinationRels.set(id, {
			id,
			type: `${relationshipNs}/${type}`,
			target: relativePartTarget(presentationPath, part),
			mode: 'Internal',
		});
		return id;
	};
	const appendEntry = (
		list: XmlElement,
		localName: string,
		relationshipId: string,
		id?: number,
	) => {
		const entry = destinationDoc.createElementNS(ns, `p:${localName}`);
		if (id !== undefined) entry.setAttribute('id', String(id));
		entry.setAttributeNS(relationshipNs, 'r:id', relationshipId);
		list.appendChild(entry);
		return entry;
	};
	for (const path of selected) {
		const entry = appendEntry(
			slideList,
			'sldId',
			register(copies.get(path)!, 'slide'),
			allocateSlideId(),
		);
		slideList.insertBefore(entry, existingSlides[insertAt] ?? null);
	}
	let masterList = first(root, 'sldMasterIdLst', ns);
	for (const [original, copied] of copies) {
		if (masterParts.get(original) !== 'slideMaster') continue;
		if (!masterList) {
			masterList = destinationDoc.createElementNS(ns, 'p:sldMasterIdLst');
			root.insertBefore(masterList, root.firstChild);
		}
		appendEntry(masterList, 'sldMasterId', register(copied, 'slideMaster'), allocateMasterId());
		const master = parseXml(await requiredXml(zip, copied));
		for (const entry of Array.from(master.getElementsByTagNameNS(ns, 'sldLayoutId')))
			entry.setAttribute('id', String(allocateMasterId()));
		zip.file(copied, buildXml(master));
	}
	let notesMasterList = first(root, 'notesMasterIdLst', ns);
	for (const [original, copied] of copies) {
		if (masterParts.get(original) !== 'notesMaster') continue;
		if (!notesMasterList) {
			notesMasterList = destinationDoc.createElementNS(ns, 'p:notesMasterIdLst');
			root.insertBefore(notesMasterList, first(root, 'handoutMasterIdLst', ns) ?? slideList);
		}
		appendEntry(notesMasterList, 'notesMasterId', register(copied, 'notesMaster'));
	}
	zip.file(presentationPath, buildXml(destinationDoc));
	zip.file(presentationRels, buildRelationshipsXml(destinationRels));
	return {
		bytes: await zip.generateAsync({ type: 'uint8array' }),
		slidePaths: selected.map((path) => copies.get(path)!),
	};
}
async function requiredXml(zip: JSZip, path: string): Promise<string> {
	const text = await zip.file(path)?.async('string');
	if (!text) throw new Error(`Required package part is missing: ${path}`);
	return text;
}

function nativeIdAllocator(
	values: readonly (string | null)[],
	min: number,
	max: number,
	kind: string,
): () => number {
	const used = new Set<number>();
	for (const value of values) {
		const id = Number(value);
		if (!Number.isInteger(id) || id < min || id > max || used.has(id))
			throw new Error(`Destination ${kind} IDs must be unique integers between ${min} and ${max}`);
		used.add(id);
	}
	let next = min;
	return () => {
		while (used.has(next) && next <= max) next++;
		if (next > max) throw new Error(`No available ${kind} IDs remain`);
		used.add(next);
		return next++;
	};
}

/** Build a portable clipboard archive containing only selected slides and their dependencies. */
export async function createSlideClipboardPackage(
	source: ArrayBuffer | Uint8Array,
	slideIndices: readonly number[],
): Promise<Uint8Array> {
	const sourceZip = await JSZip.loadAsync(source);
	const root = parseXml(await requiredXml(sourceZip, 'ppt/presentation.xml')).documentElement;
	const size = first(root, 'sldSz', root.namespaceURI ?? '');
	const width = Number(size?.getAttribute('cx')),
		height = Number(size?.getAttribute('cy'));
	if (!width || !height) throw new Error('Source slide dimensions are missing');
	let blank: Uint8Array = new Uint8Array(await buildBlankPresentationArchive({ width, height }));
	if (root.namespaceURI === 'http://purl.oclc.org/ooxml/presentationml/main') {
		const handler = new PptxHandler();
		try {
			await handler.load(blank.buffer as ArrayBuffer);
			blank = await handler.save([], { conformance: 'strict' });
		} finally {
			handler.dispose();
		}
	}
	return (await importSlidePackage(blank, source, { slideIndices })).bytes;
}
