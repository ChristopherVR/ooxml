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
 * Other XML is parsed per call, exactly as before.
 *
 * The cache is scoped to one load: the runtime opens a reader after the
 * package is opened and closes it when the load settles or the handler is
 * disposed, so editing and save always read the package directly. Reads go
 * through the guarded entry (`opc/zip-limits`), so the uncompressed size and
 * entry limits apply unchanged.
 *
 * @module load-part-reader
 */
import type { XmlObject } from '../../types';

interface TextPartFile {
	async(kind: 'string'): Promise<string>;
}

interface TextPartStore {
	file(path: string): TextPartFile | null;
}

interface PartParser {
	parse(xml: string): unknown;
}

interface PartEntry {
	file: TextPartFile;
	text: Promise<string>;
	relationships: Map<PartParser, Promise<XmlObject | undefined>>;
}

const openReaders = new WeakMap<object, LoadPartReader>();

/** Owns package text and relationship reads for one load, never for editing or save. */
export class LoadPartReader {
	private readonly parts = new Map<string, PartEntry>();
	private closed = false;

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
		const entry: PartEntry = { file, text: file.async('string'), relationships: new Map() };
		this.parts.set(path, entry);
		// A failed read must not poison a retry or delete a newer entry.
		entry.text.catch(() => {
			if (this.parts.get(path) === entry) {
				this.parts.delete(path);
			}
		});
		return entry;
	}

	readText(path: string): Promise<string | undefined> {
		return this.entry(path)?.text ?? Promise.resolve(undefined);
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
			const pending = entry.text.then((text) => parseText(parser, text));
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
	}
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
	return zip.file(path)?.async('string') ?? Promise.resolve(undefined);
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
	const text = await zip.file(path)?.async('string');
	return text ? parseText(parser, text) : undefined;
}
