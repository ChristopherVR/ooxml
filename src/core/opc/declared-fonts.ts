/**
 * The font families an Office package declares, read from a few small parts without parsing the
 * document: Word's font table and styles, Excel's style sheet, PowerPoint's and every product's
 * themes, PowerPoint's presentation part, slide masters and first slides, and Visio's face names.
 * Each product calls this as soon as a file's bytes arrive, so its webfont stylesheet can start
 * downloading while the full parse runs (`ooxml-ui`'s `prefetchWebfonts`).
 *
 * Only small, well-known parts are inflated (`fflate`'s filter reads the central directory and
 * skips everything else), each bounded in size, so the cost is a few milliseconds even for a
 * large file. A file that is not a zip (an encrypted package, a legacy binary) gives nothing.
 * Embedded fonts are listed separately: a product never requests a family the file embeds.
 *
 * @module opc/declared-fonts
 */
import { strFromU8, unzipSync } from 'fflate';

export interface DeclaredFonts {
	/** Families the package names, in first-seen order, theme references (`+mn-lt`) excluded. */
	families: string[];
	/** Families the package embeds (PowerPoint `p:embeddedFont`, Word `w:embed*`). */
	embedded: string[];
}

/** Parts worth reading, by path. Slides are capped: the first ones and the masters set the faces. */
const PARTS: readonly RegExp[] = [
	/^word\/(fontTable|styles)\.xml$/u,
	/^xl\/styles\.xml$/u,
	/^ppt\/presentation\.xml$/u,
	/^ppt\/slideMasters\/slideMaster\d+\.xml$/u,
	/^ppt\/slides\/slide([1-9]|1[0-9]|20)\.xml$/u,
	/^visio\/document\.xml$/u,
	/^(word|xl|ppt|visio)\/theme\/theme\d+\.xml$/u,
];

/** Largest part inflated (bytes); a bigger one is skipped rather than slowing the open. */
const MAX_PART_BYTES = 2 * 1024 * 1024;
/** Every part together. */
const MAX_TOTAL_BYTES = 8 * 1024 * 1024;

/** Attribute patterns that name a family: DrawingML, WordprocessingML, SpreadsheetML, Visio. */
const FAMILY_PATTERNS: readonly RegExp[] = [
	/\btypeface="([^"]+)"/gu,
	/<w:font\s+w:name="([^"]+)"/gu,
	/\bw:(?:ascii|hAnsi|eastAsia|cs)="([^"]+)"/gu,
	/<name\s+val="([^"]+)"/gu,
	/<FaceName\b[^>]*?\bNameU="([^"]+)"/gu,
];

const decodeXml = (value: string) =>
	value
		.replace(/&quot;/gu, '"')
		.replace(/&apos;/gu, "'")
		.replace(/&lt;/gu, '<')
		.replace(/&gt;/gu, '>')
		.replace(/&amp;/gu, '&');

function add(into: Map<string, string>, raw: string): void {
	const family = decodeXml(raw).trim();
	// `+mn-lt` and friends are theme references; empty and absurd names are noise.
	if (!family || family.startsWith('+') || family.length > 64) return;
	const key = family.toLowerCase();
	if (!into.has(key)) into.set(key, family);
}

/** Embedded families: PowerPoint's `p:embeddedFont`, Word's fonts with `w:embed*` children. */
function collectEmbedded(name: string, xml: string, into: Map<string, string>): void {
	if (name === 'ppt/presentation.xml') {
		for (const match of xml.matchAll(/<p:embeddedFont>\s*<p:font\b[^>]*?\btypeface="([^"]+)"/gu)) {
			if (match[1]) add(into, match[1]);
		}
	}
	if (name === 'word/fontTable.xml') {
		for (const match of xml.matchAll(/<w:font\s+w:name="([^"]+)"[^>]*>([\s\S]*?)<\/w:font>/gu)) {
			if (match[1] && /<w:embed(Regular|Bold|Italic|BoldItalic)\b/u.test(match[2] ?? '')) {
				add(into, match[1]);
			}
		}
	}
}

/** The families `bytes` (a .docx, .xlsx, .pptx or .vsdx and their template variants) declares. */
export function readDeclaredFonts(bytes: Uint8Array | ArrayBuffer): DeclaredFonts {
	const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
	let total = 0;
	let parts: Record<string, Uint8Array>;
	try {
		parts = unzipSync(data, {
			filter: (file) => {
				if (!PARTS.some((pattern) => pattern.test(file.name))) return false;
				if (file.originalSize > MAX_PART_BYTES || total + file.originalSize > MAX_TOTAL_BYTES) {
					return false;
				}
				total += file.originalSize;
				return true;
			},
		});
	} catch {
		return { families: [], embedded: [] };
	}
	const families = new Map<string, string>();
	const embedded = new Map<string, string>();
	// Fixed order (theme first), so the result does not depend on the zip's entry order.
	for (const name of Object.keys(parts).sort(byPriority)) {
		const part = parts[name];
		if (!part) continue;
		const xml = strFromU8(part);
		for (const pattern of FAMILY_PATTERNS) {
			for (const match of xml.matchAll(pattern)) if (match[1]) add(families, match[1]);
		}
		collectEmbedded(name, xml, embedded);
	}
	const embeddedKeys = new Set(embedded.keys());
	return {
		families: [...families].filter(([key]) => !embeddedKeys.has(key)).map(([, name]) => name),
		embedded: [...embedded.values()],
	};
}

function priority(name: string): number {
	if (name.includes('/theme/')) return 0;
	if (/styles\.xml$|fontTable\.xml$|presentation\.xml$|document\.xml$/u.test(name)) return 1;
	if (name.includes('slideMaster')) return 2;
	return 3;
}

function byPriority(a: string, b: string): number {
	return priority(a) - priority(b) || a.localeCompare(b, 'en', { numeric: true });
}
