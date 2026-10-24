import { XmlObject } from '../../types';
import type { PptxThemePartSchemes } from '../../types';
import { PptxHandlerRuntime as PptxHandlerRuntimeBase } from './PptxHandlerRuntimeThemeRefResolution';
import { colorSchemeFromMap, fontSchemeFromMap } from './theme-part-model';
import type { ThemePartModel } from './theme-part-model';
import { THEME_COLOR_SCHEME_KEYS } from './theme-scheme-edit';

/** The font map prefix of each `a:fontScheme` group. */
const FONT_SLOTS = [
	['mj', 'a:majorFont'],
	['mn', 'a:minorFont'],
] as const;

/**
 * Reading one theme part (`ppt/theme/themeN.xml`).
 *
 * {@link parseThemePartXml} turns a part's XML into a {@link ThemePartModel} and changes nothing;
 * {@link cacheThemePartModel} records what the save pipeline and font resolution need of it, keyed
 * by the part's path. The load pipeline does both for every part a master uses
 * (`parseThemePart`); {@link getThemePartSchemes} only reads.
 */
export class PptxHandlerRuntime extends PptxHandlerRuntimeBase {
	/** Parse one theme part's XML. Side-effect free. */
	protected parseThemePartXml(themeXml: string): ThemePartModel {
		const themeData = this.parser.parse(themeXml) as XmlObject;
		const themeRoot = themeData['a:theme'] as XmlObject | undefined;
		const themeElements = themeRoot?.['a:themeElements'] as XmlObject | undefined;
		const colorScheme = themeElements?.['a:clrScheme'] as XmlObject | undefined;
		const fontScheme = themeElements?.['a:fontScheme'] as XmlObject | undefined;
		const fmtScheme = themeElements?.['a:fmtScheme'] as XmlObject | undefined;
		const text = (value: unknown): string | undefined => String(value || '').trim() || undefined;

		const colorMap: Record<string, string> = { ...this.getDefaultSchemeColorMap() };
		for (const key of THEME_COLOR_SCHEME_KEYS) {
			const parsed = this.parseColorChoice(colorScheme?.[`a:${key}`] as XmlObject | undefined);
			if (parsed) {
				colorMap[key] = parsed;
			}
		}
		// Aliases (`tx1`/`bg1`/`tx2`/`bg2`) are resolved LAZILY through the active master's
		// `<p:clrMap>` at colour-lookup time (see `resolveThemeColor`). Baking them as `tx1 = dk1`
		// here would leak the default mapping even when a master remaps them (dark themes that swap
		// `bg1` and `tx1`). Phase 2 Stream B / C-H4.

		const fontMap: Record<string, string> = {};
		const typeface = (node: unknown): string | undefined =>
			this.normalizeTypefaceToken(String((node as XmlObject | undefined)?.['@_typeface'] || ''));
		for (const [prefix, tag] of FONT_SLOTS) {
			const group = fontScheme?.[tag] as XmlObject | undefined;
			const latin = typeface(group?.['a:latin']);
			if (latin) {
				fontMap[`${prefix}-lt`] = latin;
				fontMap[`${prefix}-ea`] = latin;
				fontMap[`${prefix}-cs`] = latin;
			}
			// The common Office layout has empty `typeface=""` placeholders, which must not stomp
			// the latin fallback.
			const eastAsia = typeface(group?.['a:ea']);
			if (eastAsia) {
				fontMap[`${prefix}-ea`] = eastAsia;
			}
			const complexScript = typeface(group?.['a:cs']);
			if (complexScript) {
				fontMap[`${prefix}-cs`] = complexScript;
			}
		}

		const objectDefaultsNode = themeRoot?.['a:objectDefaults'] as XmlObject | undefined;
		const objectDefaults = objectDefaultsNode && {
			spDef: objectDefaultsNode['a:spDef'] as unknown,
			lnDef: objectDefaultsNode['a:lnDef'] as unknown,
			txDef: objectDefaultsNode['a:txDef'] as unknown,
		};
		const hasObjectDefaults =
			objectDefaults !== undefined &&
			(objectDefaults.spDef !== undefined ||
				objectDefaults.lnDef !== undefined ||
				objectDefaults.txDef !== undefined);

		return {
			name: text(themeRoot?.['@_name']),
			colorSchemeName: text(colorScheme?.['@_name']),
			fontSchemeName: text(fontScheme?.['@_name']),
			colorMap,
			fontMap,
			// Per-script fonts (`<a:font script="Hans|Hant|Arab|...">`): M4.
			majorFontScripts: this.collectFontScriptOverrides(
				fontScheme?.['a:majorFont'] as XmlObject | undefined,
			),
			minorFontScripts: this.collectFontScriptOverrides(
				fontScheme?.['a:minorFont'] as XmlObject | undefined,
			),
			// objectDefaults: M5.
			objectDefaults: hasObjectDefaults ? objectDefaults : undefined,
			extraClrSchemeLst: themeRoot?.['a:extraClrSchemeLst'] as unknown,
			custClrLst: themeRoot?.['a:custClrLst'] as unknown,
			extLst: themeRoot?.['a:extLst'] as unknown,
			formatScheme: fmtScheme ? this.parseFormatScheme(fmtScheme) : undefined,
		};
	}

	/**
	 * Record a parsed part under its path: its original XML (byte-stable passthrough on save,
	 * C-H3), its names, per-script fonts and the raw subtrees the save re-emits.
	 */
	protected cacheThemePartModel(themePath: string, themeXml: string, model: ThemePartModel): void {
		this.originalThemeXmlByPath.set(themePath, themeXml);
		const keyed: ReadonlyArray<readonly [Map<string, unknown>, unknown]> = [
			[this.masterThemeNames, model.name],
			[this.masterThemeColorSchemeNames, model.colorSchemeName],
			[this.masterThemeFontSchemeNames, model.fontSchemeName],
			[this.masterThemeObjectDefaults, model.objectDefaults],
			[this.masterThemeExtraClrSchemeLst, model.extraClrSchemeLst],
			[this.masterThemeCustClrLst, model.custClrLst],
			[this.masterThemeExtLst, model.extLst],
		];
		for (const [map, value] of keyed) {
			if (value !== undefined) {
				map.set(themePath, value);
			}
		}
		if (Object.keys(model.majorFontScripts).length > 0) {
			this.masterThemeMajorFontScripts.set(themePath, model.majorFontScripts);
		}
		if (Object.keys(model.minorFontScripts).length > 0) {
			this.masterThemeMinorFontScripts.set(themePath, model.minorFontScripts);
		}
	}

	/**
	 * Parse `<a:font script="..." typeface="..."/>` children of a major or minor font node into a
	 * `script -> typeface` dictionary. fast-xml-parser collapses repeated tags into arrays, so the
	 * array form is walked regardless of how many siblings are present. Phase 4 Stream A / M4.
	 */
	protected collectFontScriptOverrides(fontNode: XmlObject | undefined): Record<string, string> {
		const overrides: Record<string, string> = {};
		if (!fontNode) {
			return overrides;
		}
		const fontEntries = this.ensureArray(fontNode['a:font']) as XmlObject[];
		for (const entry of fontEntries) {
			const script = String(entry?.['@_script'] || '').trim();
			const typeface = this.normalizeTypefaceToken(String(entry?.['@_typeface'] || ''));
			if (script && typeface) {
				overrides[script] = typeface;
			}
		}
		return overrides;
	}

	/**
	 * The colour scheme, font scheme and name of one theme part in the package, as the part holds
	 * them now (a scheme edit already written to it included). Reads only: the loaded theme, the
	 * masters and what a save writes are unchanged. `undefined` when the package has no such part.
	 */
	public async getThemePartSchemes(themePath: string): Promise<PptxThemePartSchemes | undefined> {
		const path = themePath.trim().replace(/\\/g, '/');
		const themeXml = await this.zip.file(path)?.async('string');
		if (themeXml === undefined) {
			return undefined;
		}
		const model = this.parseThemePartXml(themeXml);
		return {
			path,
			...(model.name ? { name: model.name } : {}),
			colorScheme: colorSchemeFromMap(model.colorMap),
			fontScheme: fontSchemeFromMap(model.fontMap, model.majorFontScripts, model.minorFontScripts),
		};
	}
}
