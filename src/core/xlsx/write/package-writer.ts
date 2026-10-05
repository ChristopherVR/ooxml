import JSZip from 'jszip';
import {
	buildContentTypesXml,
	buildRelationshipsXml,
	relationshipsPartFor,
	type RelationshipInput,
} from '../../opc/index.js';
import { CONTENT_TYPES, type SourceIndex } from '../read/package.js';

/** Content types by extension for binary parts that may lack a source declaration. */
const EXTENSION_TYPES: Record<string, string> = {
	png: 'image/png',
	jpeg: 'image/jpeg',
	jpg: 'image/jpeg',
	gif: 'image/gif',
	bmp: 'image/bmp',
	tif: 'image/tiff',
	tiff: 'image/tiff',
	emf: 'image/x-emf',
	wmf: 'image/x-wmf',
	svg: 'image/svg+xml',
	webp: 'image/webp',
	bin: CONTENT_TYPES.printerSettings,
	vml: CONTENT_TYPES.vml,
	xml: CONTENT_TYPES.xml,
	rels: CONTENT_TYPES.rels,
};

/** Relationships of one part, with id allocation that never reuses a kept id. */
export class RelationshipSet {
	readonly entries = new Map<string, RelationshipInput>();

	keep(id: string, rel: RelationshipInput): void {
		this.entries.set(id, rel);
	}

	add(type: string, target: string, external = false): string {
		let n = this.entries.size + 1;
		while (this.entries.has(`rId${n}`)) n++;
		const id = `rId${n}`;
		this.entries.set(id, { type, target, mode: external ? 'External' : 'Internal' });
		return id;
	}

	/** The id of an existing relationship of `type` to `target`, or a new one. */
	ensure(type: string, target: string): string {
		for (const [id, rel] of this.entries) if (rel.type === type && rel.target === target) return id;
		return this.add(type, target);
	}
}

/** Collects the output parts, their content types and relationships, then zips them. */
export class PackageWriter {
	private readonly parts = new Map<string, Uint8Array | string>();
	private readonly types = new Map<string, string>();
	/** Output parts still holding the source bytes of a carried part (output name to source name). */
	private readonly verbatim = new Map<string, string>();

	constructor(readonly source: SourceIndex | undefined) {}

	has(partName: string): boolean {
		return this.parts.has(partName);
	}

	add(partName: string, data: Uint8Array | string, contentType?: string): void {
		this.verbatim.delete(partName);
		this.parts.set(partName, data);
		const type = contentType ?? this.source?.contentType(partName);
		if (type) this.types.set(partName, type);
	}

	/** Writes the relationships of `partName` (nothing when the set is empty). */
	rels(partName: string, set: RelationshipSet): void {
		if (!set.entries.size) return;
		this.add(
			relationshipsPartFor(partName),
			buildRelationshipsXml(set.entries),
			CONTENT_TYPES.rels,
		);
	}

	/** The first `pattern(n)` (n = 1, 2, ...) not yet written and not reserved. */
	uniqueName(pattern: (n: number) => string, reserved?: ReadonlySet<string>): string {
		for (let n = 1; ; n++) {
			const name = pattern(n);
			if (!this.parts.has(name) && !reserved?.has(name)) return name;
		}
	}

	/**
	 * Copies a source part byte for byte, with its relationships and, recursively, every
	 * internal part they reference that has not been written yet.
	 */
	carry(partName: string, as = partName): void {
		const source = this.source;
		const bytes = source?.bytes(partName);
		if (!source || !bytes || this.parts.has(as)) return;
		this.add(as, bytes, source.contentType(partName));
		this.verbatim.set(as, partName);
		const relsPart = relationshipsPartFor(partName);
		const relsBytes = source.bytes(relsPart);
		if (!relsBytes) return;
		this.add(relationshipsPartFor(as), relsBytes, CONTENT_TYPES.rels);
		for (const rel of source.rels(partName).values()) {
			const target = source.target(partName, rel);
			if (target) this.carry(target);
		}
	}

	/** Carried parts written byte for byte so far, as [output name, source name] pairs. */
	carriedParts(): [output: string, source: string][] {
		return [...this.verbatim];
	}

	/** Replaces a part's content, keeping its content type. */
	replace(partName: string, data: Uint8Array | string): void {
		if (!this.parts.has(partName)) return;
		this.verbatim.delete(partName);
		this.parts.set(partName, data);
	}

	private contentTypesXml(): string {
		const defaults = new Map<string, string>([
			['rels', CONTENT_TYPES.rels],
			['xml', CONTENT_TYPES.xml],
		]);
		const overrides = new Map<string, string>();
		const sourceDefaults = this.source?.contentTypes.defaults;
		for (const name of [...this.parts.keys()].sort()) {
			if (name === '[Content_Types].xml') continue;
			const ext = name.split('.').pop()?.toLowerCase() ?? '';
			const type =
				this.types.get(name) ??
				sourceDefaults?.get(ext) ??
				EXTENSION_TYPES[ext] ??
				'application/octet-stream';
			if (
				!defaults.has(ext) &&
				ext !== 'xml' &&
				(sourceDefaults?.get(ext) === type || EXTENSION_TYPES[ext] === type)
			) {
				defaults.set(ext, type);
				continue;
			}
			if (defaults.get(ext) !== type) overrides.set(`/${name}`, type);
		}
		return buildContentTypesXml({ defaults, overrides });
	}

	async build(): Promise<Uint8Array> {
		const zip = new JSZip();
		const date = new Date(Date.UTC(1980, 0, 1));
		const options = { date, createFolders: false } as const;
		zip.file('[Content_Types].xml', this.contentTypesXml(), options);
		const names = [...this.parts.keys()].sort((a, b) => rank(a) - rank(b) || a.localeCompare(b));
		for (const name of names) zip.file(name, this.parts.get(name) ?? '', options);
		return zip.generateAsync({
			type: 'uint8array',
			compression: 'DEFLATE',
			compressionOptions: { level: 6 },
		});
	}
}

const rank = (name: string): number =>
	name === '_rels/.rels'
		? 0
		: name.startsWith('docProps/')
			? 1
			: name === 'xl/workbook.xml'
				? 2
				: 3;
