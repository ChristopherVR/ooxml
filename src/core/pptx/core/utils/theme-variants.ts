/**
 * Real theme variants: the `themeVariants/` parts of a package.
 *
 * PowerPoint's own themes (`.thmx`) carry their variants as a second set of presentation parts
 * under `/themeVariants/variantN/theme/...`, listed by `/themeVariants/themeVariantManager.xml`
 * (root relationship `http://schemas.microsoft.com/office/2011/relationships/themeVariants`).
 * A deck saved by PowerPoint does not carry them (the variants live with the installed theme),
 * so most decks have none; a package that does gets those variants offered instead of computed
 * ones.
 *
 * @module utils/theme-variants
 */
import type JSZip from 'jszip';

import type {
	PptxThemeColorScheme,
	PptxThemeFontScheme,
	PptxThemeVariant,
	XmlObject,
} from '../types';
import { THEME_COLOR_SCHEME_KEYS } from '../types/theme';
import { backgroundStylePresetFor, isDarkColorMap } from '../types/background-style';

const THEME_VARIANTS_REL = 'http://schemas.microsoft.com/office/2011/relationships/themeVariants';

function asArray<T>(value: T | T[] | undefined): T[] {
	return value === undefined ? [] : Array.isArray(value) ? value : [value];
}

function dirname(path: string): string {
	return path.slice(0, Math.max(0, path.lastIndexOf('/')));
}

function stripRoot(target: string): string {
	return target.replace(/^\/+/u, '');
}

/** Resolve `target` against the folder that holds the relationships' source part. */
function resolveTarget(sourceDir: string, target: string): string {
	if (target.startsWith('/')) {
		return stripRoot(target);
	}
	const parts = sourceDir ? sourceDir.split('/') : [];
	for (const segment of target.split('/')) {
		if (segment === '..') {
			parts.pop();
		} else if (segment !== '.' && segment !== '') {
			parts.push(segment);
		}
	}
	return parts.join('/');
}

async function readXml(
	zip: JSZip,
	parse: (xml: string) => XmlObject,
	path: string,
): Promise<XmlObject | undefined> {
	const text = await zip.file(path)?.async('string');
	return text === undefined ? undefined : parse(text);
}

async function relationships(
	zip: JSZip,
	parse: (xml: string) => XmlObject,
	relsPath: string,
	sourceDir: string,
): Promise<{ id: string; type: string; target: string }[]> {
	const xml = await readXml(zip, parse, relsPath);
	const root = xml?.['Relationships'] as XmlObject | undefined;
	return asArray(root?.['Relationship'] as XmlObject | XmlObject[] | undefined).map((rel) => ({
		id: String(rel['@_Id'] ?? ''),
		type: String(rel['@_Type'] ?? ''),
		target: resolveTarget(sourceDir, String(rel['@_Target'] ?? '')),
	}));
}

function colourOf(node: XmlObject | undefined): string | undefined {
	const srgb = node?.['a:srgbClr'] as XmlObject | undefined;
	const sys = node?.['a:sysClr'] as XmlObject | undefined;
	const hex = String(srgb?.['@_val'] ?? sys?.['@_lastClr'] ?? '').trim();
	return /^[0-9a-f]{6}$/iu.test(hex) ? `#${hex.toUpperCase()}` : undefined;
}

function colorSchemeOf(scheme: XmlObject | undefined): PptxThemeColorScheme | undefined {
	if (!scheme) {
		return undefined;
	}
	const result = {} as Record<string, string>;
	for (const key of THEME_COLOR_SCHEME_KEYS) {
		const colour = colourOf(scheme[`a:${key}`] as XmlObject | undefined);
		if (!colour) {
			return undefined;
		}
		result[key] = colour;
	}
	return result as unknown as PptxThemeColorScheme;
}

function fontSchemeOf(scheme: XmlObject | undefined): PptxThemeFontScheme {
	const group = (node: XmlObject | undefined) => {
		const latin = String((node?.['a:latin'] as XmlObject | undefined)?.['@_typeface'] ?? '');
		const ea = String((node?.['a:ea'] as XmlObject | undefined)?.['@_typeface'] ?? '');
		const cs = String((node?.['a:cs'] as XmlObject | undefined)?.['@_typeface'] ?? '');
		return {
			...(latin ? { latin } : {}),
			...(ea ? { eastAsia: ea } : {}),
			...(cs ? { complexScript: cs } : {}),
		};
	};
	return {
		majorFont: group(scheme?.['a:majorFont'] as XmlObject | undefined),
		minorFont: group(scheme?.['a:minorFont'] as XmlObject | undefined),
	};
}

/** The `a:fmtScheme` element of a theme part, as the text the part carries. */
function formatSchemeText(themeXml: string): string | undefined {
	const match = /<a:fmtScheme\b[\s\S]*?<\/a:fmtScheme>/u.exec(themeXml);
	return match?.[0];
}

function clrMapOf(master: XmlObject | undefined): Record<string, string> | undefined {
	const node = (master?.['p:sldMaster'] as XmlObject | undefined)?.['p:clrMap'] as
		| XmlObject
		| undefined;
	if (!node) {
		return undefined;
	}
	return Object.fromEntries(
		Object.entries(node)
			.filter(([key]) => key.startsWith('@_'))
			.map(([key, value]) => [key.slice(2), String(value)]),
	);
}

/**
 * The variants a package carries, in the manager's order, for the deck's slide width (the first
 * size group when none matches). Empty when the package has no `themeVariants` relationship.
 */
export async function loadThemeVariants(
	zip: JSZip,
	parse: (xml: string) => XmlObject,
	deckSlideWidthEmu?: number,
): Promise<PptxThemeVariant[]> {
	const rootRels = await relationships(zip, parse, '_rels/.rels', '');
	const managerPath = rootRels.find((rel) => rel.type === THEME_VARIANTS_REL)?.target;
	if (!managerPath) {
		return [];
	}
	const manager = await readXml(zip, parse, managerPath);
	const list = (manager?.['t:themeVariantManager'] as XmlObject | undefined)?.[
		't:themeVariantLst'
	] as XmlObject | undefined;
	const entries = asArray(list?.['t:themeVariant'] as XmlObject | XmlObject[] | undefined);
	if (entries.length === 0) {
		return [];
	}
	const managerRels = await relationships(
		zip,
		parse,
		`${dirname(managerPath)}/_rels/${managerPath.slice(managerPath.lastIndexOf('/') + 1)}.rels`,
		dirname(managerPath),
	);
	const widths = entries.map((entry) => Number(entry['@_cx']));
	const width =
		deckSlideWidthEmu && widths.includes(deckSlideWidthEmu) ? deckSlideWidthEmu : widths[0];
	const variants: PptxThemeVariant[] = [];
	const seen = new Set<string>();
	for (const entry of entries) {
		if (Number(entry['@_cx']) !== width) {
			continue;
		}
		const id = String(entry['@_vid'] ?? '');
		const themeManager = managerRels.find(
			(rel) => rel.id === String(entry['@_r:id'] ?? ''),
		)?.target;
		if (!id || seen.has(id) || !themeManager) {
			continue;
		}
		const variant = await readVariant(zip, parse, themeManager, id, String(entry['@_name'] ?? ''));
		if (variant) {
			seen.add(id);
			variants.push(variant);
		}
	}
	return variants;
}

async function readVariant(
	zip: JSZip,
	parse: (xml: string) => XmlObject,
	themeManagerPath: string,
	id: string,
	name: string,
): Promise<PptxThemeVariant | undefined> {
	const themeDir = dirname(themeManagerPath);
	const managerRels = await relationships(
		zip,
		parse,
		`${themeDir}/_rels/${themeManagerPath.slice(themeManagerPath.lastIndexOf('/') + 1)}.rels`,
		themeDir,
	);
	const themePath =
		managerRels.find((rel) => rel.type.endsWith('/relationships/theme'))?.target ??
		`${themeDir}/theme1.xml`;
	const themeText = await zip.file(themePath)?.async('string');
	if (themeText === undefined) {
		return undefined;
	}
	const elements = (parse(themeText)['a:theme'] as XmlObject | undefined)?.['a:themeElements'] as
		| XmlObject
		| undefined;
	const colorScheme = colorSchemeOf(elements?.['a:clrScheme'] as XmlObject | undefined);
	if (!colorScheme) {
		return undefined;
	}
	const formatSchemeXml = formatSchemeText(themeText);
	const masterPath = `${dirname(themeDir)}/slideMasters/slideMaster1.xml`;
	const master = await readXml(zip, parse, masterPath);
	const bgRef = (
		(master?.['p:sldMaster'] as XmlObject | undefined)?.['p:cSld'] as XmlObject | undefined
	)?.['p:bg'] as XmlObject | undefined;
	const ref = bgRef?.['p:bgRef'] as XmlObject | undefined;
	const clrMap = clrMapOf(master);
	const preset = ref
		? backgroundStylePresetFor(
				Number.parseInt(String(ref['@_idx'] ?? ''), 10),
				String((ref['a:schemeClr'] as XmlObject | undefined)?.['@_val'] ?? ''),
				isDarkColorMap(clrMap),
			)
		: undefined;
	return {
		id,
		name: name || String((parse(themeText)['a:theme'] as XmlObject)['@_name'] ?? ''),
		colorScheme,
		fontScheme: fontSchemeOf(elements?.['a:fontScheme'] as XmlObject | undefined),
		...(formatSchemeXml ? { formatSchemeXml } : {}),
		...(preset ? { backgroundStyleIndex: preset.index } : {}),
	};
}
