import JSZip from 'jszip';
import {
	RELATIONSHIP_TYPES,
	contentTypeForPart,
	parseContentTypes,
	parseRelationships,
	relationshipsPartFor,
	resolvePartPath,
	type ContentTypes,
	type Relationship,
} from '../../opc/index.js';

/** Largest accepted package (compressed). */
export const MAX_INPUT_BYTES = 50 * 1024 * 1024;
/** Largest number of zip entries. */
export const MAX_PARTS = 10_000;
/** Largest total uncompressed size, guarding against zip bombs. */
export const MAX_UNCOMPRESSED_BYTES = 300 * 1024 * 1024;

/** Content types of the SpreadsheetML parts the reader and writer know. */
export const CONTENT_TYPES = {
	workbook: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
	workbookMacro: 'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
	template: 'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml',
	templateMacro: 'application/vnd.ms-excel.template.macroEnabled.main+xml',
	worksheet: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
	chartsheet: 'application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml',
	sharedStrings: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
	styles: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
	theme: 'application/vnd.openxmlformats-officedocument.theme+xml',
	comments: 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
	threadedComments: 'application/vnd.ms-excel.threadedcomments+xml',
	persons: 'application/vnd.ms-excel.person+xml',
	table: 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
	drawing: 'application/vnd.openxmlformats-officedocument.drawing+xml',
	chart: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
	vml: 'application/vnd.openxmlformats-officedocument.vmlDrawing',
	core: 'application/vnd.openxmlformats-package.core-properties+xml',
	app: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
	rels: 'application/vnd.openxmlformats-package.relationships+xml',
	xml: 'application/xml',
	printerSettings: 'application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings',
	sheetMetadata: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml',
} as const;

/** The loaded package with cached relationship and content-type lookups. */
export class SourceIndex {
	readonly contentTypes: ContentTypes;
	private readonly relsCache = new Map<string, Map<string, Relationship>>();
	private readonly textCache = new Map<string, string>();
	private readonly decoder = new TextDecoder('utf-8');
	private readonly lowerNames = new Map<string, string>();

	constructor(readonly parts: ReadonlyMap<string, Uint8Array>) {
		for (const name of parts.keys()) this.lowerNames.set(name.toLowerCase(), name);
		this.contentTypes = parseContentTypes(this.text('[Content_Types].xml'));
	}

	/** The stored name of a part; OPC part names compare case-insensitively. */
	actual(partName: string): string {
		return this.parts.has(partName)
			? partName
			: (this.lowerNames.get(partName.toLowerCase()) ?? partName);
	}

	has(partName: string): boolean {
		return this.parts.has(this.actual(partName));
	}

	bytes(partName: string): Uint8Array | undefined {
		return this.parts.get(this.actual(partName));
	}

	/** A part decoded as UTF-8 text (BOM stripped), or `undefined` when absent. */
	text(partName: string): string | undefined {
		const cached = this.textCache.get(partName);
		if (cached !== undefined) return cached;
		const bytes = this.bytes(partName);
		if (!bytes) return undefined;
		let text = this.decoder.decode(bytes);
		if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
		this.textCache.set(partName, text);
		return text;
	}

	contentType(partName: string): string | undefined {
		return contentTypeForPart(this.contentTypes, partName);
	}

	/** Relationships declared by `partName` (`''` for the package root). */
	rels(partName: string): Map<string, Relationship> {
		const relsPart = partName === '' ? '_rels/.rels' : relationshipsPartFor(partName);
		let rels = this.relsCache.get(relsPart);
		if (!rels) {
			rels = parseRelationships(this.text(relsPart));
			this.relsCache.set(relsPart, rels);
		}
		return rels;
	}

	/** The part a relationship points at, or `undefined` for external targets. */
	target(partName: string, rel: Relationship): string | undefined {
		if (rel.mode === 'External') return undefined;
		return this.actual(resolvePartPath(partName === '' ? '/' : partName, rel.target));
	}

	/** The first internal target of a relationship type declared by `partName`. */
	targetOfType(partName: string, type: string): string | undefined {
		for (const rel of this.rels(partName).values())
			if (rel.type === type) return this.target(partName, rel);
		return undefined;
	}

	/** The main workbook part (`xl/workbook.xml` in every Excel-written file). */
	workbookPart(): string | undefined {
		const fromRels = this.targetOfType('', RELATIONSHIP_TYPES.officeDocument);
		if (fromRels && this.has(fromRels)) return fromRels;
		return this.has('xl/workbook.xml') ? 'xl/workbook.xml' : undefined;
	}
}

function uncompressedSize(entry: JSZip.JSZipObject): number {
	const data = (entry as unknown as { _data?: { uncompressedSize?: unknown } })._data;
	const size = Number(data?.uncompressedSize ?? 0);
	return Number.isFinite(size) ? size : 0;
}

/** Unzips a package into raw parts, enforcing the size and part-count limits. */
export async function readZipParts(
	input: Uint8Array | ArrayBuffer,
): Promise<Map<string, Uint8Array>> {
	const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
	if (bytes.byteLength > MAX_INPUT_BYTES)
		throw new Error('XLSX package exceeds the 50 MiB compressed input limit');
	let zip: JSZip;
	try {
		zip = await JSZip.loadAsync(bytes);
	} catch (error) {
		throw new Error(
			`Not a valid XLSX package: ${error instanceof Error ? error.message : String(error)}`,
		);
	}
	const entries = Object.values(zip.files).filter((entry) => !entry.dir);
	if (entries.length > MAX_PARTS) throw new Error('XLSX package exceeds the 10,000 part limit');
	const total = entries.reduce((sum, entry) => sum + uncompressedSize(entry), 0);
	if (total > MAX_UNCOMPRESSED_BYTES)
		throw new Error('XLSX package exceeds the 300 MiB uncompressed content limit');
	const parts = new Map<string, Uint8Array>();
	let read = 0;
	for (const entry of entries) {
		const data = await entry.async('uint8array');
		read += data.byteLength;
		if (read > MAX_UNCOMPRESSED_BYTES)
			throw new Error('XLSX package exceeds the 300 MiB uncompressed content limit');
		parts.set(entry.name.replace(/^\//, ''), data);
	}
	return parts;
}
