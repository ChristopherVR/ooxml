/**
 * Shares package part reads between the loaders that run during one
 * `PptxHandler.load()`.
 *
 * A load reads the same parts many times: `ppt/_rels/presentation.xml.rels`,
 * every master and layout relationship part and the slide XML are each read
 * by several independent loaders, and the chart label preflight walks every
 * relationship part once per chart. Each read inflates the entry again.
 *
 * While a reader is open for a package, {@link readPartText} returns one
 * shared text promise per part and {@link readPartXml} parses each
 * relationship part once per parser. Callers receive a fresh deep copy of a
 * relationship tree, so a caller that mutates it cannot affect another.
 * Chart readers share the input DOM and keep each model's compatibility tree independent.
 * Templates and themes opt in to sharing a parse, with independent trees for mutable models.
 *
 * The cache is scoped to one load: the runtime opens a reader after the
 * package is opened and closes it when the load settles or the handler is
 * disposed, so editing and save always read the package directly. Reads go
 * through the guarded entry (`opc/zip-limits`), so the uncompressed size and
 * entry limits apply unchanged.
 *
 * @module load-part-reader
 */
import type { ChartXmlSource } from '../../../../chart/keep-source';
import { annotateCustomGeometryCommandOrder } from '../../geometry/custom-geometry-command-order';
import { annotateParagraphSiblingOrder } from './paragraph-sibling-order';
import { annotateSmartArtTextOrder } from './smartart-text-order';
import { canonicalRelationshipsXml } from '../../../../opc/relationships';
import { decodeXmlBytes } from '../../../../xml/index';
import type { XmlObject } from '../../types';
import { readOriginalChartPart } from '../../utils/chart-part-repair';
import { projectChartCompatibilityTree } from '../../utils/chart-compat-tree';

interface TextPartFile {
	async(kind: 'string'): Promise<string>;
	async(kind: 'uint8array'): Promise<Uint8Array>;
}

/**
 * A part's text. ECMA-376 Part 2 allows UTF-8 and UTF-16; UTF-16 read as UTF-8 starts with NUL
 * characters or byte order mark replacements, so only such a part is read again as bytes.
 */
async function textOf(file: TextPartFile): Promise<string> {
	const text = await file.async('string');
	if (!/^(?:\uFFFD\uFFFD|[^\0]?\0)/u.test(text)) return text;
	return decodeXmlBytes(await file.async('uint8array'), { fatal: false });
}

interface TextPartStore {
	file(path: string): TextPartFile | null;
}

interface PartParser {
	parse(xml: string): unknown;
}

interface GalleryTree {
	parser: PartParser;
	source: string;
	xml: XmlObject;
}

interface PartEntry {
	galleryTree?: GalleryTree;
	file: TextPartFile;
	text: Promise<string>;
	relationships: Map<PartParser, Promise<XmlObject | undefined>>;
	chartXml?: { parser: PartParser; parsed: Promise<ChartPartXml | undefined> };
}

interface ChartPartXml {
	text: string;
	original: ChartXmlSource | undefined;
	compatibility(): XmlObject;
}

function parseChartXml(parser: PartParser, text: string): ChartPartXml | undefined {
	if (!text) return undefined;
	const original = readOriginalChartPart(text);
	let fallback: XmlObject | undefined;
	return {
		text,
		original,
		compatibility() {
			const projected = projectChartCompatibilityTree(parser, text, original);
			if (projected) return projected;
			// The chart readers only inspect the tree. Cloning on a second reference keeps
			// any raw nodes retained by the resulting models independent.
			if (fallback) return copyParsedXmlWithOrder(fallback, text);
			fallback = parser.parse(text) as XmlObject;
			return fallback;
		},
	};
}

const openReaders = new WeakMap<object, LoadPartReader>();

/** Owns package reads for one load, never for editing or save. */
export class LoadPartReader {
	private readonly parts = new Map<string, PartEntry>();
	private readonly derived = new Map<string, Promise<unknown>>();
	private closed = false;
	private strictRelationships = false;

	constructor(readonly zip: TextPartStore) {
		openReaders.set(zip, this);
	}

	private entry(path: string): PartEntry | undefined {
		const file = this.zip.file(path);
		if (!file) {
			this.parts.delete(path);
			return undefined;
		}
		const previous = this.parts.get(path);
		// JSZip replaces the entry object on every write, so a write or removal
		// made during the load is never served from a stale read.
		if (previous?.file === file) {
			return previous;
		}
		const entry: PartEntry = { file, text: textOf(file), relationships: new Map() };
		this.parts.set(path, entry);
		// A failed read must not poison a retry or delete a newer entry.
		entry.text.catch(() => {
			if (this.parts.get(path) === entry) {
				this.parts.delete(path);
			}
		});
		return entry;
	}

	/**
	 * From now on relationship parts reach the parser as the strict reader saw them, so a signed
	 * package loads exactly the relationships its verifier read.
	 */
	useStrictRelationships(): void {
		this.strictRelationships = true;
		this.parts.clear();
		this.derived.clear();
	}

	/** `compute()`, run once per key while the load lasts (a failure is not kept). */
	derive<T>(key: string, compute: () => Promise<T>): Promise<T> {
		let value = this.derived.get(key) as Promise<T> | undefined;
		if (!value) {
			const pending = compute();
			this.derived.set(key, pending);
			pending.catch(() => {
				if (this.derived.get(key) === pending) {
					this.derived.delete(key);
				}
			});
			value = pending;
		}
		return value;
	}

	readText(path: string): Promise<string | undefined> {
		return this.entry(path)?.text ?? Promise.resolve(undefined);
	}

	readChartXml(parser: PartParser, path: string): Promise<ChartPartXml | undefined> {
		const entry = this.entry(path);
		if (!entry) return Promise.resolve(undefined);
		if (entry.chartXml?.parser === parser) return entry.chartXml.parsed;
		const parsed = entry.text.then((text) => parseChartXml(parser, text));
		entry.chartXml = { parser, parsed };
		return parsed;
	}

	/** Snapshot a template's source tree for the independent gallery model. */
	rememberGalleryTree(path: string, parser: PartParser, source: string, xml: XmlObject): void {
		const entry = this.parts.get(path);
		if (!entry || entry.file !== this.zip.file(path)) return;
		entry.galleryTree = { parser, source, xml: copyParsedXmlWithOrder(xml, source) };
	}

	takeGalleryTree(path: string, parser: PartParser, source: string): XmlObject | undefined {
		const entry = this.parts.get(path);
		const gallery = entry?.galleryTree;
		if (!entry || !gallery || entry.file !== this.zip.file(path)) return undefined;
		delete entry.galleryTree;
		return gallery.parser === parser && gallery.source === source ? gallery.xml : undefined;
	}

	async readXml(parser: PartParser, path: string): Promise<XmlObject | undefined> {
		const entry = this.entry(path);
		if (!entry) {
			return undefined;
		}
		if (!path.endsWith('.rels')) {
			return parseText(parser, await entry.text);
		}
		// Parser identity includes the Strict-normalizing wrapper, so a parse made
		// before Strict detection is not reused after that wrapper is installed.
		let parsed = entry.relationships.get(parser);
		if (!parsed) {
			const pending = entry.text.then((text) =>
				parseText(parser, this.strictRelationships ? canonicalRelationshipsXml(text) : text),
			);
			entry.relationships.set(parser, pending);
			pending.catch(() => {
				if (entry.relationships.get(parser) === pending) {
					entry.relationships.delete(parser);
				}
			});
			parsed = pending;
		}
		return structuredClone(await parsed);
	}

	/** Whether this reader still serves reads for its package. */
	get isOpen(): boolean {
		return !this.closed;
	}

	/** Drops every cached read; later reads of the package go to the archive. */
	close(): void {
		this.closed = true;
		if (openReaders.get(this.zip) === this) {
			openReaders.delete(this.zip);
		}
		this.parts.clear();
		this.derived.clear();
	}
}

/** Copies a tree and restores the sibling-order annotations stored outside the tree. */
export function copyParsedXmlWithOrder(tree: XmlObject, source: string): XmlObject {
	const copy = structuredClone(tree);
	// The OMML order markers are own properties and are already present.
	if (source.includes('custGeom')) annotateCustomGeometryCommandOrder(source, copy);
	if (source.includes('dataModel') || source.includes('txBody'))
		annotateSmartArtTextOrder(source, copy);
	if (source.includes('<a:p')) annotateParagraphSiblingOrder(source, copy);
	return copy;
}

function parseText(parser: PartParser, text: string): XmlObject | undefined {
	return text ? (parser.parse(text) as XmlObject) : undefined;
}

/**
 * The text of a package part, or `undefined` when the package has no such
 * part. Shared with the open load's reader when there is one.
 */
export function readPartText(zip: TextPartStore, path: string): Promise<string | undefined> {
	const reader = openReaders.get(zip);
	if (reader) {
		return reader.readText(path);
	}
	const file = zip.file(path);
	return file ? textOf(file) : Promise.resolve(undefined);
}

/** Share chart XML input within a load. Every chart model gets its own compatibility tree. */
export async function readChartXml(
	zip: TextPartStore,
	parser: PartParser,
	path: string,
): Promise<ChartPartXml | undefined> {
	const reader = openReaders.get(zip);
	if (reader) return reader.readChartXml(parser, path);
	const file = zip.file(path);
	return file ? parseChartXml(parser, await textOf(file)) : undefined;
}

/**
 * A value derived from package parts, computed once while a load is open for `zip` (every chart of
 * the load shares it) and afresh outside a load, where the package may have changed.
 */
export function deriveForLoad<T>(
	zip: TextPartStore,
	key: string,
	compute: () => Promise<T>,
): Promise<T> {
	const reader = openReaders.get(zip);
	return reader ? reader.derive(key, compute) : compute();
}

/**
 * A parsed package part, or `undefined` when the part is missing or empty.
 * Relationship parts are parsed once per open load and returned as copies.
 */
export async function readPartXml(
	zip: TextPartStore,
	parser: PartParser,
	path: string,
): Promise<XmlObject | undefined> {
	const reader = openReaders.get(zip);
	if (reader) {
		return reader.readXml(parser, path);
	}
	const file = zip.file(path);
	const text = file ? await textOf(file) : undefined;
	return text ? parseText(parser, text) : undefined;
}
