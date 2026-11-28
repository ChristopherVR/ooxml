import { children, first, getR, getW, parseXml } from './xml';

export interface FontFallback {
	alternate?: string;
	generic?: 'serif' | 'sans-serif' | 'monospace' | 'cursive' | 'fantasy';
}

const families: ReadonlyMap<string, NonNullable<FontFallback['generic']>> = new Map([
	['roman', 'serif'],
	['swiss', 'sans-serif'],
	// Word also assigns modern to proportional fonts such as Montserrat.
	['modern', 'sans-serif'],
	['script', 'cursive'],
	['decorative', 'fantasy'],
]);

/** Font substitutions declared in the document, keyed by case-insensitive font name. */
export function parseFontTable(source: string): ReadonlyMap<string, FontFallback> {
	const result = new Map<string, FontFallback>();
	for (const font of children(parseXml(source).documentElement, 'font')) {
		const name = getW(font, 'name')?.trim();
		if (!name) continue;
		const alternate = getW(first(font, 'altName'), 'val')?.trim();
		const pitch = getW(first(font, 'pitch'), 'val');
		const generic =
			pitch === 'fixed' ? 'monospace' : families.get(getW(first(font, 'family'), 'val') ?? '');
		if (alternate || generic)
			result.set(name.toLowerCase(), {
				...(alternate ? { alternate } : {}),
				...(generic ? { generic } : {}),
			});
	}
	return result;
}

/** A font the package embeds (`w:embedRegular` and its siblings), as `word/fontTable.xml` lists it. */
export interface EmbeddedFontRef {
	family: string;
	bold: boolean;
	italic: boolean;
	/** The obfuscated font part, e.g. `word/fonts/font1.odttf`. */
	partName: string;
	/** The `w:fontKey` GUID the first 32 bytes of the part are XORed with (ISO/IEC 29500-2 §14.2.1). */
	fontKey: string;
}

const EMBEDS = [
	['embedRegular', false, false],
	['embedBold', true, false],
	['embedItalic', false, true],
	['embedBoldItalic', true, true],
] as const;

/**
 * The embedded faces of the font table. `target` resolves a relationship id of the font table
 * to the part it points at (the font table's own relationships, resolved against `word/`).
 */
export function parseEmbeddedFonts(
	source: string,
	target: (relationshipId: string) => string | undefined,
): EmbeddedFontRef[] {
	const result: EmbeddedFontRef[] = [];
	for (const font of children(parseXml(source).documentElement, 'font')) {
		const family = getW(font, 'name')?.trim();
		if (!family) continue;
		for (const [local, bold, italic] of EMBEDS) {
			const embed = first(font, local);
			const id = getR(embed, 'id');
			const fontKey = getW(embed, 'fontKey');
			const partName = id ? target(id) : undefined;
			if (partName && fontKey) result.push({ family, bold, italic, partName, fontKey });
		}
	}
	return result;
}
