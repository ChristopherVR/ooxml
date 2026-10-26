/**
 * Embedded font variant decoding, shared by the eager and the deferred load
 * paths so both produce byte-identical `rawFontData` and `dataUrl` values.
 *
 * A variant's package part is either an EOT container (Google Slides and
 * some other producers write MicroType Express compressed EOT `.fntdata`
 * parts) or an OOXML XOR-obfuscated font (ECMA-376 Part 2, 14.2.1). Turning
 * either into a usable font means decompressing or de-obfuscating it and
 * base64-encoding the result, which dominates the load time of decks that
 * embed several faces. {@link createDeferredEmbeddedFont} keeps the package
 * bytes and does that work only when a variant's data is first read.
 */
import type { PptxEmbeddedFont } from '../types';
import { extractFontFromEot, isEotFormat, parseEotHeader } from './eot-parser';
import { deobfuscateFont, detectFontFormat, extractGuidFromPartName } from './font-deobfuscation';

/** Where an embedded variant's part sits and how it is declared. */
export interface EmbeddedFontPartSource {
	/** Typeface name from `p:font/@typeface`. */
	name: string;
	bold: boolean;
	italic: boolean;
	/** Relationship id from the variant's `r:id`. */
	originalRId: string;
	/** Package path of the font part. */
	partPath: string;
	/** The part's bytes as stored in the package. */
	originalPartBytes: Uint8Array;
	/** The variant's `fontKey` attribute, when declared. */
	fontKey?: string | undefined;
}

/** The fields decoding produces. */
interface DecodedFields {
	dataUrl: string;
	format: PptxEmbeddedFont['format'];
	rawFontData: Uint8Array | undefined;
}

/**
 * The obfuscation GUID for a non-EOT part: the `fontKey` attribute first,
 * then a GUID in the part name. `undefined` for EOT parts and parts with
 * neither.
 */
export function resolveEmbeddedFontGuid(source: EmbeddedFontPartSource): string | undefined {
	if (isEotFormat(source.originalPartBytes)) {
		return undefined;
	}
	const { fontKey } = source;
	const fromKey = fontKey
		? (extractGuidFromPartName(fontKey) ?? fontKey.replace(/[{}]/g, '').trim())
		: null;
	return fromKey || extractGuidFromPartName(source.partPath) || undefined;
}

/** True when `bytes` start with a TrueType, OpenType, TTC or WOFF/WOFF2 signature. */
function hasSfntSignature(bytes: Uint8Array): boolean {
	if (bytes.length < 4) {
		return false;
	}
	const [v0, v1, v2, v3] = bytes;
	return (
		(v0 === 0x00 && v1 === 0x01 && v2 === 0x00 && v3 === 0x00) ||
		(v0 === 0x4f && v1 === 0x54 && v2 === 0x54 && v3 === 0x4f) ||
		(v0 === 0x74 && v1 === 0x74 && v2 === 0x63 && v3 === 0x66) ||
		(v0 === 0x77 && v1 === 0x4f && v2 === 0x46 && (v3 === 0x46 || v3 === 0x32))
	);
}

function warnInvalidSignature(name: string, bytes: Uint8Array): void {
	const hex = Array.from(bytes.subarray(0, 4), (b) => b.toString(16)).join(',');
	console.warn(
		`[pptx-viewer] Embedded font "${name}" has invalid sfnt header [${hex}] ` +
			`after processing; skipping (web font fallback will be used)`,
	);
}

/**
 * Checks that need no decompression and no full copy of the part: a non-EOT
 * part needs a GUID, and its de-obfuscated header must be a font signature.
 * EOT parts pass; their payload can only be checked by decompressing it.
 * Logs the same diagnostics as a failed decode.
 */
export function isEmbeddedFontPartPlausible(source: EmbeddedFontPartSource): boolean {
	const bytes = source.originalPartBytes;
	if (bytes.length === 0) {
		return false;
	}
	if (isEotFormat(bytes)) {
		return true;
	}
	const guid = resolveEmbeddedFontGuid(source);
	if (!guid) {
		console.warn(
			`[pptx-viewer] Embedded font "${source.name}" at "${source.partPath}" has no fontKey attribute and no GUID in filename; skipping`,
		);
		return false;
	}
	// The XOR layer only covers the first 32 bytes, so the header is enough.
	const header = deobfuscateFont(bytes.subarray(0, Math.min(bytes.length, 32)), guid);
	if (header.length < 4) {
		return false;
	}
	if (!hasSfntSignature(header)) {
		warnInvalidSignature(source.name, header);
		return false;
	}
	return true;
}

/** Base64-encode `bytes` in chunks. */
function uint8ArrayToBase64(bytes: Uint8Array): string {
	let binary = '';
	const chunkSize = 8192;
	for (let i = 0; i < bytes.length; i += chunkSize) {
		const chunk = bytes.subarray(i, Math.min(i + chunkSize, bytes.length));
		for (let j = 0; j < chunk.length; j++) {
			binary += String.fromCharCode(chunk[j]);
		}
	}
	return btoa(binary);
}

/**
 * Decompress or de-obfuscate one variant, validate its signature and build
 * its data URL. Returns `null` (after logging why) when the part cannot be
 * turned into a usable font.
 */
export function decodeEmbeddedFontPart(source: EmbeddedFontPartSource): PptxEmbeddedFont | null {
	const { name, partPath, originalPartBytes } = source;
	try {
		let fontData: Uint8Array;
		const fontGuid = resolveEmbeddedFontGuid(source);
		if (isEotFormat(originalPartBytes)) {
			const extracted = extractFontFromEot(originalPartBytes);
			if (!extracted) {
				const header = parseEotHeader(originalPartBytes);
				console.info(
					`[pptx-viewer] Embedded font "${name}" uses EOT format${header?.isCompressed ? ' with MTX compression' : ''}; extraction failed, will use web font fallback`,
				);
				return null;
			}
			fontData = extracted.fontData;
		} else if (fontGuid) {
			fontData = deobfuscateFont(originalPartBytes, fontGuid);
		} else {
			console.warn(
				`[pptx-viewer] Embedded font "${name}" at "${partPath}" has no fontKey attribute and no GUID in filename; skipping`,
			);
			return null;
		}
		if (fontData.length < 4) {
			return null;
		}
		if (!hasSfntSignature(fontData)) {
			warnInvalidSignature(name, fontData);
			return null;
		}
		const format = detectFontFormat(fontData);
		const mimeType =
			format === 'woff2'
				? 'font/woff2'
				: format === 'woff'
					? 'font/woff'
					: format === 'opentype'
						? 'font/otf'
						: 'font/ttf';
		return {
			name,
			dataUrl: `data:${mimeType};base64,${uint8ArrayToBase64(fontData)}`,
			bold: source.bold,
			italic: source.italic,
			format,
			rawFontData: fontData,
			partPath,
			fontGuid,
			originalRId: source.originalRId,
			originalPartBytes,
		};
	} catch {
		return null;
	}
}

/**
 * A variant whose `dataUrl`, `format` and `rawFontData` are decoded on first
 * read, or on `load()`. Metadata reads (`name`, `bold`, `italic`,
 * `partPath`, `fontGuid`, `originalRId`, `originalPartBytes`) never decode.
 *
 * The descriptor owns its bytes, so it stays usable after the handler that
 * produced it is disposed or loads another deck. Its enumerable fields and
 * their order match an eagerly decoded variant, so spreading, cloning and
 * JSON serialisation give the same result (and decode it). `load` and
 * `isLoaded` are non-enumerable: they are runtime capabilities, not model
 * data. A part that fails to decode keeps its metadata and has an empty
 * `dataUrl` and no `rawFontData`, the same as an absent font to renderers.
 */
export function createDeferredEmbeddedFont(source: EmbeddedFontPartSource): PptxEmbeddedFont {
	const fontGuid = resolveEmbeddedFontGuid(source);
	let fields: DecodedFields | undefined;
	let pending: Promise<PptxEmbeddedFont | null> | undefined;
	const decode = (): DecodedFields => {
		if (!fields) {
			const decoded = decodeEmbeddedFontPart(source);
			fields = {
				dataUrl: decoded?.dataUrl ?? '',
				format: decoded?.format,
				rawFontData: decoded?.rawFontData,
			};
		}
		return fields;
	};
	const font: PptxEmbeddedFont = {
		name: source.name,
		get dataUrl() {
			return decode().dataUrl;
		},
		set dataUrl(value) {
			decode().dataUrl = value;
		},
		bold: source.bold,
		italic: source.italic,
		get format() {
			return decode().format;
		},
		set format(value) {
			decode().format = value;
		},
		get rawFontData() {
			return decode().rawFontData;
		},
		set rawFontData(value) {
			decode().rawFontData = value;
		},
		partPath: source.partPath,
		fontGuid,
		originalRId: source.originalRId,
		originalPartBytes: source.originalPartBytes,
	};
	Object.defineProperties(font, {
		load: {
			value: () =>
				(pending ??= Promise.resolve().then(() =>
					decode().rawFontData?.length || decode().dataUrl ? font : null,
				)),
		},
		isLoaded: { get: () => fields !== undefined },
	});
	return font;
}

/**
 * True when saving should write a deferred variant's original part bytes
 * under its original part and relationship: it was never decoded (reading
 * its data now would only decode it to re-encode the same bytes), or it
 * failed to decode. Never decodes. Fonts that are not deferred, and
 * decoded ones, keep the `rawFontData` save path.
 */
export function shouldSaveEmbeddedFontVerbatim(font: PptxEmbeddedFont): boolean {
	if (font.isLoaded === undefined || !font.originalPartBytes?.length) {
		return false;
	}
	if (!font.originalRId || !font.partPath) {
		return false;
	}
	return font.isLoaded === false || !font.rawFontData?.length;
}
