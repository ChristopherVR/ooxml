import { XmlObject } from '../../types';
import type { PptxThemePartSchemes } from '../../types';
import { PptxHandlerRuntime as PptxHandlerRuntimeBase } from './PptxHandlerRuntimeThemeRefResolution';
import { parseTheme } from '../../../../drawingml/theme';
import { NS, parseXml, type XmlElement } from '../../../../xml/index';
import type { DrawingTheme } from '../../../../drawingml/theme-model';
import {
	colorSchemeFromMap,
	fontSchemeFromMap,
	themeSchemesFromDrawingTheme,
} from './theme-part-model';
import type { ThemePartModel } from './theme-part-model';
import { themeFormatSchemeView } from './theme-format-scheme-view';

/**
 * Reading one theme part (`ppt/theme/themeN.xml`).
 *
 * {@link parseThemePartXml} turns a part's XML into a {@link ThemePartModel} and changes nothing;
 * {@link cacheThemePartModel} records what the save pipeline and font resolution need of it, keyed
 * by the part's path. The load pipeline does both for every part a master uses
 * (`parseThemePart`); {@link getThemePartSchemes} only reads.
 */
export class PptxHandlerRuntime extends PptxHandlerRuntimeBase {
	/**
	 * Parse one theme part's XML. Side-effect free. The names and the colour, font and format
	 * schemes are read by the shared DrawingML reader (`parseTheme`); the subtrees the save
	 * re-emits verbatim come from the object tree.
	 */
	protected parseThemePartXml(
		themeXml: string,
		themeData = this.parser.parse(themeXml) as XmlObject,
		drawingRoot: XmlElement | null = this.readDrawingThemeRoot(themeXml) ?? null,
	): ThemePartModel {
		const themeRoot = themeData['a:theme'] as XmlObject | undefined;
		const themeElements = themeRoot?.['a:themeElements'] as XmlObject | undefined;
		const fmtScheme = themeElements?.['a:fmtScheme'] as XmlObject | undefined;
		// Aliases (`tx1`/`bg1`/`tx2`/`bg2`) are not baked into the colour map: they resolve LAZILY
		// through the active master's `<p:clrMap>` at colour-lookup time (see `resolveThemeColor`),
		// since a master may remap them (dark themes that swap `bg1` and `tx1`).
		const drawingTheme = this.readDrawingTheme(drawingRoot);
		const schemes = themeSchemesFromDrawingTheme(drawingTheme, this.getDefaultSchemeColorMap());

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
			...schemes,
			// objectDefaults: M5.
			objectDefaults: hasObjectDefaults ? objectDefaults : undefined,
			extraClrSchemeLst: themeRoot?.['a:extraClrSchemeLst'] as unknown,
			custClrLst: themeRoot?.['a:custClrLst'] as unknown,
			extLst: themeRoot?.['a:extLst'] as unknown,
			formatScheme: !fmtScheme
				? undefined
				: drawingTheme.formatScheme
					? themeFormatSchemeView(
							drawingTheme.formatScheme,
							fmtScheme,
							this.themeStyleViewContext(),
						)
					: this.parseFormatScheme(fmtScheme),
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
	 * The shared reader's view of a theme part. A part whose root is not `a:theme`, or that the DOM
	 * parser rejects (not namespace well-formed), reads as an empty theme, so its slots keep the
	 * schema defaults.
	 */
	protected readDrawingThemeRoot(themeXml: string): XmlElement | undefined {
		try {
			const root = parseXml(themeXml, { label: 'PPTX theme' }).documentElement;
			if (root.localName === 'theme' && root.namespaceURI === NS.a) return root;
		} catch {
			// Fall through to the empty theme.
		}
		return undefined;
	}

	private readDrawingTheme(root: XmlElement | null): DrawingTheme {
		try {
			if (root) return parseTheme(root);
		} catch {
			// Fall through to the empty theme.
		}
		return {
			colorScheme: { colors: {} },
			fontScheme: { major: { scripts: {} }, minor: { scripts: {} } },
		};
	}

	/**
	 * The colour, font and format schemes and name of one theme part in the package, as the part holds
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
			...(model.formatScheme ? { formatScheme: model.formatScheme } : {}),
		};
	}
}
