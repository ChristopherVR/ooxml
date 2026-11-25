import { children, first, getW, parseXml } from './xml';

export interface FontFallback {
	alternate?: string;
	generic?: 'serif' | 'sans-serif' | 'monospace' | 'cursive' | 'fantasy';
}

const families: ReadonlyMap<string, NonNullable<FontFallback['generic']>> = new Map([
	['roman', 'serif'],
	['swiss', 'sans-serif'],
	['modern', 'monospace'],
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
		const generic = families.get(getW(first(font, 'family'), 'val') ?? '');
		if (alternate || generic)
			result.set(name.toLowerCase(), {
				...(alternate ? { alternate } : {}),
				...(generic ? { generic } : {}),
			});
	}
	return result;
}
