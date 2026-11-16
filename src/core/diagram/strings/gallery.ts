// The Change Colors and SmartArt Styles gallery names (titles, sections and Office's 38 colour and
// 14 quick-style definitions, `DIAGRAM_NATIVE_*_ENTRIES`) in every product locale, so Word and
// Excel show the names PowerPoint shows. The translations are PowerPoint's own gallery dictionary
// (`pptx.gallery.smartArt*` in the pptx locales package), expanded here into English-keyed names.
// They were written by agents and are not verified against Office's localized gallery names.
import { DIAGRAM_NATIVE_COLOR_ENTRIES, DIAGRAM_NATIVE_STYLE_ENTRIES } from '../catalogue-native';

export const galleryEn = {
	'Change Colors': 'Change Colors',
	'SmartArt Styles': 'SmartArt Styles',
	'Primary Theme Colors': 'Primary Theme Colors',
	Colorful: 'Colorful',
	'Accent 1': 'Accent 1',
	'Accent 2': 'Accent 2',
	'Accent 3': 'Accent 3',
	'Accent 4': 'Accent 4',
	'Accent 5': 'Accent 5',
	'Accent 6': 'Accent 6',
	'Best Match for Document': 'Best Match for Document',
	'3-D': '3-D',
	'Dark 1 Outline': 'Dark 1 Outline',
	'Dark 2 Outline': 'Dark 2 Outline',
	'Dark 2 Fill': 'Dark 2 Fill',
	'Colorful - Accent Colors': 'Colorful - Accent Colors',
	'Colorful Range - Accent Colors 2 to 3': 'Colorful Range - Accent Colors 2 to 3',
	'Colorful Range - Accent Colors 3 to 4': 'Colorful Range - Accent Colors 3 to 4',
	'Colorful Range - Accent Colors 4 to 5': 'Colorful Range - Accent Colors 4 to 5',
	'Colorful Range - Accent Colors 5 to 6': 'Colorful Range - Accent Colors 5 to 6',
	'Colored Outline - Accent 1': 'Colored Outline - Accent 1',
	'Colored Fill - Accent 1': 'Colored Fill - Accent 1',
	'Gradient Range - Accent 1': 'Gradient Range - Accent 1',
	'Gradient Loop - Accent 1': 'Gradient Loop - Accent 1',
	'Transparent Gradient Range - Accent 1': 'Transparent Gradient Range - Accent 1',
	'Colored Outline - Accent 2': 'Colored Outline - Accent 2',
	'Colored Fill - Accent 2': 'Colored Fill - Accent 2',
	'Gradient Range - Accent 2': 'Gradient Range - Accent 2',
	'Gradient Loop - Accent 2': 'Gradient Loop - Accent 2',
	'Transparent Gradient Range - Accent 2': 'Transparent Gradient Range - Accent 2',
	'Colored Outline - Accent 3': 'Colored Outline - Accent 3',
	'Colored Fill - Accent 3': 'Colored Fill - Accent 3',
	'Gradient Range - Accent 3': 'Gradient Range - Accent 3',
	'Gradient Loop - Accent 3': 'Gradient Loop - Accent 3',
	'Transparent Gradient Range - Accent 3': 'Transparent Gradient Range - Accent 3',
	'Colored Outline - Accent 4': 'Colored Outline - Accent 4',
	'Colored Fill - Accent 4': 'Colored Fill - Accent 4',
	'Gradient Range - Accent 4': 'Gradient Range - Accent 4',
	'Gradient Loop - Accent 4': 'Gradient Loop - Accent 4',
	'Transparent Gradient Range - Accent 4': 'Transparent Gradient Range - Accent 4',
	'Colored Outline - Accent 5': 'Colored Outline - Accent 5',
	'Colored Fill - Accent 5': 'Colored Fill - Accent 5',
	'Gradient Range - Accent 5': 'Gradient Range - Accent 5',
	'Gradient Loop - Accent 5': 'Gradient Loop - Accent 5',
	'Transparent Gradient Range - Accent 5': 'Transparent Gradient Range - Accent 5',
	'Colored Outline - Accent 6': 'Colored Outline - Accent 6',
	'Colored Fill - Accent 6': 'Colored Fill - Accent 6',
	'Gradient Range - Accent 6': 'Gradient Range - Accent 6',
	'Gradient Loop - Accent 6': 'Gradient Loop - Accent 6',
	'Transparent Gradient Range - Accent 6': 'Transparent Gradient Range - Accent 6',
	'Simple Fill': 'Simple Fill',
	'White Outline': 'White Outline',
	'Subtle Effect': 'Subtle Effect',
	'Moderate Effect': 'Moderate Effect',
	'Intense Effect': 'Intense Effect',
	Polished: 'Polished',
	Inset: 'Inset',
	Cartoon: 'Cartoon',
	Powder: 'Powder',
	'Brick Scene': 'Brick Scene',
	'Flat Scene': 'Flat Scene',
	'Metallic Scene': 'Metallic Scene',
	'Sunset Scene': 'Sunset Scene',
	"Bird's Eye Scene": "Bird's Eye Scene",
} as const;

export type GalleryStrings = Record<keyof typeof galleryEn, string>;

const STYLE_IDS = [
	'simple1',
	'simple2',
	'simple3',
	'simple4',
	'simple5',
	'3d1',
	'3d2',
	'3d3',
	'3d4',
	'3d5',
	'3d6',
	'3d7',
	'3d8',
	'3d9',
] as const;

/** One locale's gallery dictionary; `{accent}`, `{n}`, `{from}` and `{to}` are placeholders. */
export interface GalleryTemplates {
	changeColors: string;
	smartArtStyles: string;
	mainScheme: string;
	colorful: string;
	/** A section heading, `Accent {n}`. */
	accent: string;
	bestMatch: string;
	threeD: string;
	dark1Outline: string;
	dark2Outline: string;
	dark2Fill: string;
	colorful1: string;
	colorfulRange: string;
	coloredOutline: string;
	coloredFill: string;
	gradientRange: string;
	gradientLoop: string;
	transparentGradientRange: string;
	styles: Record<(typeof STYLE_IDS)[number], string>;
}

const fill = (template: string, values: Record<string, number>): string =>
	template.replace(/\{(\w+)\}/gu, (match, name: string) => String(values[name] ?? match));

const COLOR_VARIANTS = [
	'coloredOutline',
	'coloredFill',
	'gradientRange',
	'gradientLoop',
	'transparentGradientRange',
] as const;
const MAIN_SCHEME: Readonly<Record<string, 'dark1Outline' | 'dark2Outline' | 'dark2Fill'>> = {
	accent0_1: 'dark1Outline',
	accent0_2: 'dark2Outline',
	accent0_3: 'dark2Fill',
};

function colorName(id: string, t: GalleryTemplates): string | undefined {
	const main = MAIN_SCHEME[id];
	if (main) return t[main];
	if (id === 'colorful1') return t.colorful1;
	const range = /^colorful(\d)$/u.exec(id);
	if (range) return fill(t.colorfulRange, { from: Number(range[1]), to: Number(range[1]) + 1 });
	const accent = /^accent(\d)_(\d)$/u.exec(id);
	const variant = accent ? COLOR_VARIANTS[Number(accent[2]) - 1] : undefined;
	return accent && variant ? fill(t[variant], { accent: Number(accent[1]) }) : undefined;
}

const isStyleId = (id: string): id is (typeof STYLE_IDS)[number] =>
	(STYLE_IDS as readonly string[]).includes(id);

/** A locale's gallery names, keyed by the English name each replaces. */
export function galleryStrings(t: GalleryTemplates): GalleryStrings {
	const out: Record<string, string> = {
		'Change Colors': t.changeColors,
		'SmartArt Styles': t.smartArtStyles,
		'Primary Theme Colors': t.mainScheme,
		Colorful: t.colorful,
		'Best Match for Document': t.bestMatch,
		'3-D': t.threeD,
	};
	for (let n = 1; n <= 6; n++) out[`Accent ${n}`] = fill(t.accent, { n });
	for (const entry of DIAGRAM_NATIVE_COLOR_ENTRIES) {
		const name = colorName(entry.id, t);
		if (name) out[entry.name] = name;
	}
	for (const entry of DIAGRAM_NATIVE_STYLE_ENTRIES)
		if (isStyleId(entry.id)) out[entry.name] = t.styles[entry.id];
	return out as GalleryStrings;
}
