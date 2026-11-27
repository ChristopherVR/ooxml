/**
 * The fonts a page already holds locally, as files a PDF can embed: every `@font-face` rule in
 * the document's `<style>` elements whose source is a `data:` or `blob:` URL (the deck's own
 * embedded fonts, which the viewers register that way). Remote sources are skipped, so reading
 * them never makes a network request. Whether a font may be embedded is decided later by
 * `PdfFontSet` from the font's own licence flags.
 */
import type { PdfFontFile } from 'ooxml-core/pdf/write';

/** The subset of `Document` this needs, so tests can pass a fake. */
export interface FontFaceSourceDocument {
	querySelectorAll(selector: 'style'): ArrayLike<{ textContent: string | null }>;
}

/** One `@font-face` rule with a local source. */
export interface LocalFontFace {
	family: string;
	bold: boolean;
	italic: boolean;
	url: string;
}

const MAX_FACES = 200;
/** Fonts larger than this are not read (bounds memory on a hostile page). */
const MAX_FONT_BYTES = 32 * 1024 * 1024;

const unquote = (value: string) => value.trim().replace(/^(['"])(.*)\1$/s, '$2');

/**
 * One declaration's value, read up to the `;` that ends it: a `data:` URL holds `;` inside its
 * quotes and parentheses (`data:font/ttf;base64,...`), so a plain split would cut it short.
 */
function readDeclaration(body: string, name: string): string {
	const start = new RegExp(`(?:^|[;{\\s])${name}\\s*:`, 'i').exec(body);
	if (!start) return '';
	let depth = 0;
	let quote = '';
	let index = start.index + start[0].length;
	for (; index < body.length; index++) {
		const ch = body[index]!;
		if (quote) {
			if (ch === quote) quote = '';
		} else if (ch === '"' || ch === "'") quote = ch;
		else if (ch === '(') depth++;
		else if (ch === ')') depth--;
		else if (ch === ';' && depth <= 0) break;
	}
	return body.slice(start.index + start[0].length, index).trim();
}

/** Parse the `@font-face` rules of `css` that name a `data:` or `blob:` source. */
export function parseLocalFontFaces(css: string): LocalFontFace[] {
	const faces: LocalFontFace[] = [];
	for (const match of css.matchAll(/@font-face\s*\{([^}]*)\}/g)) {
		if (faces.length >= MAX_FACES) break;
		const body = match[1]!;
		const read = (name: string) =>
			new RegExp(`(?:^|[;{\\s])${name}\\s*:\\s*([^;]*)`, 'i').exec(body)?.[1]?.trim();
		const family = unquote(read('font-family') ?? '');
		const src = readDeclaration(body, 'src');
		const url = [...src.matchAll(/url\(\s*(['"]?)([^'")]+)\1\s*\)/g)]
			.map((m) => m[2]!)
			.find((candidate) => /^(data:|blob:)/i.test(candidate));
		if (!family || !url) continue;
		const weight = (read('font-weight') ?? 'normal').toLowerCase();
		const style = (read('font-style') ?? 'normal').toLowerCase();
		faces.push({
			family,
			bold: weight === 'bold' || Number.parseInt(weight, 10) >= 600,
			italic: style.startsWith('italic') || style.startsWith('oblique'),
			url,
		});
	}
	return faces;
}

function decodeDataUrl(url: string): Uint8Array | undefined {
	const comma = url.indexOf(',');
	if (comma < 0) return undefined;
	const meta = url.slice(5, comma);
	const payload = url.slice(comma + 1);
	if (!/;base64$/i.test(meta)) return undefined;
	if ((payload.length * 3) / 4 > MAX_FONT_BYTES) return undefined;
	try {
		const binary = atob(payload);
		const bytes = new Uint8Array(binary.length);
		for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
		return bytes;
	} catch {
		return undefined;
	}
}

async function readFaceBytes(url: string): Promise<Uint8Array | undefined> {
	if (/^data:/i.test(url)) return decodeDataUrl(url);
	try {
		// A blob: URL is an object in this page's memory; fetching it never leaves the device.
		const response = await fetch(url);
		const buffer = await response.arrayBuffer();
		return buffer.byteLength > MAX_FONT_BYTES ? undefined : new Uint8Array(buffer);
	} catch {
		return undefined;
	}
}

/** The local font files of `doc`'s `@font-face` rules (see the module comment). */
export async function collectLocalFontFiles(doc: FontFaceSourceDocument): Promise<PdfFontFile[]> {
	const styles = doc.querySelectorAll('style');
	let css = '';
	for (let i = 0; i < styles.length; i++) {
		const text = styles[i]!.textContent ?? '';
		if (text.includes('@font-face')) css += `${text}\n`;
	}
	const files: PdfFontFile[] = [];
	const seen = new Set<string>();
	for (const face of parseLocalFontFaces(css)) {
		const key = `${face.family.toLowerCase()}|${face.bold}|${face.italic}`;
		if (seen.has(key)) continue;
		seen.add(key);
		const data = await readFaceBytes(face.url);
		if (data) files.push({ family: face.family, bold: face.bold, italic: face.italic, data });
	}
	return files;
}
