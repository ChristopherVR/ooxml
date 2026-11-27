/**
 * Deterministic OPC packaging for the sample generators: the same parts always give the same
 * bytes, on any machine and in any time zone.
 */
import JSZip from 'jszip';

/** JSZip writes the DOS timestamp from the UTC fields, so this date is zone independent. */
const ZIP_DATE = new Date(Date.UTC(2026, 0, 15, 9, 0, 0));
/** Formats that already compress their payload are stored. */
const STORED = /\.(?:png|jpe?g|gif|webp)$/iu;

export type PackageParts = Record<string, string | Uint8Array>;

async function write(zip: JSZip): Promise<Uint8Array> {
	return zip.generateAsync({
		type: 'uint8array',
		compression: 'DEFLATE',
		compressionOptions: { level: 9 },
		platform: 'DOS',
	});
}

function add(zip: JSZip, name: string, data: string | Uint8Array): void {
	zip.file(name, data, {
		date: ZIP_DATE,
		createFolders: false,
		compression: STORED.test(name) ? 'STORE' : 'DEFLATE',
	});
}

/** A package from its parts, in the order given (`[Content_Types].xml` conventionally first). */
export async function buildPackage(parts: PackageParts): Promise<Uint8Array> {
	const zip = new JSZip();
	for (const [name, data] of Object.entries(parts)) add(zip, name, data);
	return write(zip);
}

/** The parts of a package as bytes, in archive order. */
export async function readPackage(bytes: Uint8Array): Promise<Map<string, Uint8Array>> {
	const zip = await JSZip.loadAsync(bytes);
	const parts = new Map<string, Uint8Array>();
	for (const entry of Object.values(zip.files)) {
		if (!entry.dir) parts.set(entry.name, await entry.async('uint8array'));
	}
	return parts;
}

/**
 * Rewrites a package some writer produced so its bytes depend on its parts alone: fixed entry
 * dates, one compression setting, no folder entries. Part order and content are kept.
 */
export async function normalizePackage(bytes: Uint8Array): Promise<Uint8Array> {
	const zip = new JSZip();
	for (const [name, data] of await readPackage(bytes)) add(zip, name, data);
	return write(zip);
}

const decoder = new TextDecoder();
/** One part as text. */
export async function readPartText(bytes: Uint8Array, name: string): Promise<string | undefined> {
	const data = (await readPackage(bytes)).get(name);
	return data ? decoder.decode(data) : undefined;
}

/**
 * A copy of the package with parts added or replaced (`undefined` removes one). Replaced parts
 * keep their position; new parts go to the end.
 */
export async function patchPackage(
	bytes: Uint8Array,
	changes: Record<string, string | Uint8Array | undefined>,
): Promise<Uint8Array> {
	const ordered: PackageParts = {};
	for (const [name, data] of await readPackage(bytes)) {
		if (!(name in changes)) ordered[name] = data;
		else if (changes[name] !== undefined) ordered[name] = changes[name];
	}
	for (const [name, data] of Object.entries(changes))
		if (data !== undefined && !(name in ordered)) ordered[name] = data;
	return buildPackage(ordered);
}

/** XML escaping for element content and attribute values. */
export function xmlEscape(text: string): string {
	return text
		.replaceAll('&', '&amp;')
		.replaceAll('<', '&lt;')
		.replaceAll('>', '&gt;')
		.replaceAll('"', '&quot;');
}
