/**
 * The font files a host supplies for a PDF export, and which of them writes and measures a
 * family the document asks for: its own file, else the host's mapping, else the first supplied
 * metric-compatible or usual substitute (Carlito for Calibri, Liberation Serif for Times New
 * Roman...), as Word falls back when a font is not installed. DOM-free.
 */
import type { FontFallback } from '../font-table';
import type { PdfFontFile } from '../../pdf/write/pdf-embedded-font';
import {
	getSubstituteFonts,
	METRIC_FONT_SUBSTITUTES,
	parseTrueType,
	type TrueTypeFont,
} from '../../text/index';

/** Where the file that stands in for a family came from. */
export type DocxPdfFontSource = 'supplied' | 'mapping' | 'substitute';

export interface DocxPdfFontMatch {
	/** The family of the supplied file that writes and measures the text. */
	family: string;
	source: DocxPdfFontSource;
}

/** A parsed supplied file. */
export interface FontFace {
	family: string;
	bold: boolean;
	italic: boolean;
	font: TrueTypeFont;
}

const GENERIC = new Set(['serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui']);
const key = (family: string) =>
	family
		.trim()
		.replace(/^['"]|['"]$/g, '')
		.toLowerCase();

/** The supplied fonts, parsed once, and the family each document family resolves to. */
export class DocxPdfFontFiles {
	readonly faces: FontFace[];
	readonly #byFamily = new Map<string, FontFace[]>();
	readonly #mapping: ReadonlyMap<string, string>;
	readonly #matches = new Map<string, DocxPdfFontMatch | undefined>();

	constructor(files: readonly PdfFontFile[], substitutions: Readonly<Record<string, string>> = {}) {
		this.faces = files.flatMap((file) => {
			const font = parseTrueType(file.data);
			return font ? [{ family: file.family, bold: !!file.bold, italic: !!file.italic, font }] : [];
		});
		for (const face of this.faces) {
			const list = this.#byFamily.get(key(face.family)) ?? [];
			list.push(face);
			this.#byFamily.set(key(face.family), list);
		}
		this.#mapping = new Map(
			Object.entries(substitutions).map(([from, to]) => [key(from), key(to)] as const),
		);
	}

	/** Which supplied family writes `family` (the first of a CSS list that resolves), if any. */
	match(family: string, fallback?: FontFallback): DocxPdfFontMatch | undefined {
		const cacheKey = `${family}\u0000${fallback?.alternate ?? ''}`;
		if (this.#matches.has(cacheKey)) return this.#matches.get(cacheKey);
		const found = this.#resolve(family, fallback);
		this.#matches.set(cacheKey, found);
		return found;
	}

	#resolve(family: string, fallback?: FontFallback): DocxPdfFontMatch | undefined {
		const names = family.split(',').map(key).filter(Boolean);
		const supplied = (name: string) => this.#byFamily.get(name)?.[0]?.family;
		for (const name of names) {
			const own = supplied(name);
			if (own) return { family: own, source: 'supplied' };
		}
		for (const name of names) {
			const mapped = this.#mapping.get(name);
			const file = mapped ? supplied(mapped) : undefined;
			if (file) return { family: file, source: 'mapping' };
		}
		const alternates = fallback?.alternate ? [key(fallback.alternate)] : [];
		for (const name of [...names, ...alternates]) {
			const candidates = [
				...(Object.hasOwn(METRIC_FONT_SUBSTITUTES, name)
					? METRIC_FONT_SUBSTITUTES[name]!.fallbacks
					: []),
				...getSubstituteFonts(name),
			];
			for (const candidate of candidates) {
				if (GENERIC.has(key(candidate))) continue;
				const file = supplied(key(candidate));
				if (file) return { family: file, source: 'substitute' };
			}
		}
		return undefined;
	}

	/** The supplied face for `family` in the closest style (regular stands in for bold). */
	face(family: string, bold: boolean, italic: boolean): FontFace | undefined {
		const faces = this.#byFamily.get(key(family));
		if (!faces) return undefined;
		const score = (face: FontFace) =>
			(face.bold === bold ? 2 : 0) + (face.italic === italic ? 1 : 0);
		return faces.reduce((best, face) => (score(face) > score(best) ? face : best));
	}
}
