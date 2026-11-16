/**
 * @fileoverview Save-side theme.xml writer.
 *
 * Phase 4 Stream A / C-H3.
 *
 * Round-trip strategy:
 *
 * 1. On parse, the original `theme*.xml` body is captured into
 *    {@link PptxHandlerRuntime#originalThemeXmlByPath}. The full set of
 *    fields a PPTX theme can carry — `bgFillStyleLst`, `fillStyleLst`,
 *    `lnStyleLst`, `effectStyleLst`, `objectDefaults`, `extLst`,
 *    `extraClrSchemeLst`, `custClrLst`, … — is large; the in-memory
 *    model (`themeColorMap`, `themeFontMap`, `themeFormatScheme`) only
 *    captures a subset.
 *
 * 2. On save, if the theme has been marked dirty
 *    ({@link PptxHandlerRuntime#dirtyThemePaths}), the shared DrawingML
 *    theme writer (`drawingThemeXml`) regenerates the colour and font
 *    schemes from in-memory state, preserving the
 *    captured raw subtrees for the parts the model doesn't own
 *    (objectDefaults, extraClrSchemeLst, fmtScheme details, …).
 *    Otherwise the original XML is re-emitted verbatim, which guarantees
 *    byte-stable round-trip in the common (no-mutation) case.
 *
 * The writer hooks into `PptxHandlerRuntimeSavePipeline.save` between
 * the master / layout persistence and the presentation persistence.
 */

import {
	drawingThemeFormatSchemeXml,
	type ThemeFormatScheme,
} from '../../../../drawingml/format-scheme';
import {
	THEME_COLOR_SLOTS,
	type DrawingTheme,
	type ThemeColorScheme,
	type ThemeFontCollection,
} from '../../../../drawingml/theme-model';
import type { DrawingColor, DrawingLine } from '../../../../drawingml/types';
import { drawingThemeXml } from '../../../../drawingml/write-theme';
import { XmlObject } from '../../types';
import { PptxHandlerRuntime as PptxHandlerRuntimeBase } from './PptxHandlerRuntimeSaveSlideWriter';

/**
 * The twelve scheme slots of a colour map as `srgbClr` colours: the in-memory map is hex-typed, so
 * a slot the original part wrote as `sysClr` is written as its resolved value; a missing or
 * malformed slot is written black.
 */
function themeSlotColors(colorMap: Record<string, string>): ThemeColorScheme['colors'] {
	const colors: ThemeColorScheme['colors'] = {};
	for (const slot of THEME_COLOR_SLOTS) {
		const hex = String(colorMap[slot] || '').replace(/^#/, '');
		const value = hex.length === 6 ? hex.toUpperCase() : '000000';
		colors[slot] = { kind: 'srgb', value, transforms: [] };
	}
	return colors;
}

/** The major (`mj`) or minor (`mn`) fonts of a font map with their per-script faces. */
function themeFontCollection(
	fontMap: Record<string, string>,
	prefix: 'mj' | 'mn',
	scripts: Record<string, string> | undefined,
): ThemeFontCollection {
	return {
		latin: fontMap[`${prefix}-lt`] || 'Calibri',
		eastAsia: fontMap[`${prefix}-ea`] || '',
		complexScript: fontMap[`${prefix}-cs`] || '',
		scripts: { ...scripts },
	};
}

const PLACEHOLDER: DrawingColor = { kind: 'scheme', value: 'phClr', transforms: [] };

/**
 * The smallest schema-valid `a:fmtScheme` (CT_FillStyleList and its siblings require at least
 * three entries), for a theme part that has none: placeholder-coloured solid fills and outlines of
 * 0.5, 1 and 1.5 pt, and empty effect lists.
 */
export function minimalThemeFormatScheme(name: string): ThemeFormatScheme {
	const fill = () => ({ fill: { kind: 'solid' as const, color: structuredClone(PLACEHOLDER) } });
	const line = (widthEmu: number): DrawingLine => ({
		widthEmu,
		cap: 'flat',
		compound: 'sng',
		alignment: 'ctr',
		fill: { kind: 'solid', color: structuredClone(PLACEHOLDER) },
		dash: 'solid',
		join: { kind: 'miter', limit: 800000 },
	});
	return {
		name,
		fillStyles: [fill(), fill(), fill()],
		lineStyles: [line(6350), line(12700), line(19050)],
		effectStyles: [0, 1, 2].map(() => ({ effects: { effects: [] } })),
		backgroundFillStyles: [fill(), fill(), fill()],
	};
}

/**
 * Public mutation API for theme state. Marking a theme dirty causes the
 * save pipeline to regenerate the theme part instead of re-emitting the
 * original XML.
 */
export class PptxHandlerRuntime extends PptxHandlerRuntimeBase {
	/**
	 * Mark a theme path as dirty so the save pipeline will regenerate
	 * the theme XML from in-memory state. Optional — without this the
	 * original XML is preserved verbatim on save (C-H3).
	 */
	public markThemeDirty(themePath: string): void {
		this.dirtyThemePaths.add(themePath);
	}

	/**
	 * Mark all known theme paths dirty in one call.
	 */
	public markAllThemesDirty(): void {
		for (const themePath of this.originalThemeXmlByPath.keys()) {
			this.dirtyThemePaths.add(themePath);
		}
		for (const themePath of this.masterThemePaths.values()) {
			this.dirtyThemePaths.add(themePath);
		}
	}

	/**
	 * Persist all theme parts during save. Called from the save pipeline
	 * after master / layout XML have been flushed and before
	 * presentation.xml is serialized.
	 *
	 * Order of operations per theme path:
	 *
	 * 1. If the path is *not* in {@link dirtyThemePaths}, the existing
	 *    ZIP entry is already correct — no-op. (Original XML was placed
	 *    into the ZIP at load time.)
	 * 2. If the path is dirty, build a fresh `<a:theme>` document from
	 *    in-memory state and the captured raw subtrees, then overwrite
	 *    the ZIP entry.
	 */
	protected async persistThemeParts(): Promise<void> {
		// Build a list of (masterPath, themePath) pairs to persist. A master
		// without an entry in masterThemePaths inherited the global theme
		// path; in that case the per-master maps are empty and the global
		// themeColorMap / themeFontMap / themeFormatScheme are the source.
		const seenThemePaths = new Set<string>();

		for (const [masterPath, themePath] of this.masterThemePaths.entries()) {
			if (!themePath) {
				continue;
			}
			seenThemePaths.add(themePath);
			if (!this.dirtyThemePaths.has(themePath)) {
				continue;
			}
			const themeXml = this.buildThemeXml(themePath, masterPath);
			if (themeXml) {
				this.zip.file(themePath, themeXml);
			}
		}

		// Catch any theme paths that have an original captured but no master
		// pointing at them (e.g. orphan themes in the ZIP). Re-emit dirty ones
		// from the global state.
		for (const [themePath] of this.originalThemeXmlByPath.entries()) {
			if (seenThemePaths.has(themePath)) {
				continue;
			}
			if (!this.dirtyThemePaths.has(themePath)) {
				continue;
			}
			const themeXml = this.buildThemeXml(themePath, undefined);
			if (themeXml) {
				this.zip.file(themePath, themeXml);
			}
		}
	}

	/**
	 * Build a complete `<a:theme>` XML document from in-memory state.
	 * Returns the serialized XML string (with XML prolog), or `undefined`
	 * if there is no source data to emit.
	 *
	 * The colour and font schemes are written by the shared DrawingML theme
	 * writer (`drawingThemeXml`) from the per-master maps (or the global
	 * fallback), each slot as `srgbClr`. The authored `a:fmtScheme` and the
	 * objectDefaults / extraClrSchemeLst / custClrLst / extLst subtrees are
	 * re-emitted as captured; a part with no authored format scheme gets the
	 * shared writer's minimal one ({@link minimalThemeFormatScheme}).
	 */
	protected buildThemeXml(themePath: string, masterPath: string | undefined): string | undefined {
		const colorMap =
			(masterPath && this.masterThemeColorMaps.get(masterPath)) ||
			this.globalThemeColorMapSnapshot ||
			this.themeColorMap;
		const fontMap =
			(masterPath && this.masterThemeFontMaps.get(masterPath)) ||
			this.globalThemeFontMapSnapshot ||
			this.themeFontMap;

		const themeName = this.masterThemeNames.get(themePath) || 'Office Theme';
		const theme: DrawingTheme = {
			name: themeName,
			colorScheme: {
				name: this.masterThemeColorSchemeNames.get(themePath) || themeName,
				colors: themeSlotColors(colorMap),
			},
			fontScheme: {
				name: this.masterThemeFontSchemeNames.get(themePath) || themeName,
				major: themeFontCollection(fontMap, 'mj', this.masterThemeMajorFontScripts.get(themePath)),
				minor: themeFontCollection(fontMap, 'mn', this.masterThemeMinorFontScripts.get(themePath)),
			},
		};

		let themeRoot: XmlObject | undefined;
		try {
			const xml = drawingThemeXml(theme, {
				formatScheme: drawingThemeFormatSchemeXml(minimalThemeFormatScheme(themeName)),
			});
			themeRoot = (this.parser.parse(xml) as XmlObject)['a:theme'] as XmlObject | undefined;
		} catch (error) {
			console.warn(`Failed to build theme XML for ${themePath}:`, error);
			return undefined;
		}
		const themeElements = themeRoot?.['a:themeElements'] as XmlObject | undefined;
		if (!themeRoot || !themeElements) {
			return undefined;
		}

		// Format scheme: the original raw subtree (kept via originalThemeXmlByPath), verbatim.
		const fmtScheme = this.extractRawSubtreeFromOriginal(themePath, [
			'a:theme',
			'a:themeElements',
			'a:fmtScheme',
		]);
		if (fmtScheme !== undefined) {
			themeElements['a:fmtScheme'] = fmtScheme as XmlObject;
		}

		// objectDefaults - M5.
		const objectDefaults = this.masterThemeObjectDefaults.get(themePath);
		if (objectDefaults && (objectDefaults.spDef || objectDefaults.lnDef || objectDefaults.txDef)) {
			const od: XmlObject = {};
			if (objectDefaults.spDef !== undefined) {
				od['a:spDef'] = objectDefaults.spDef as XmlObject;
			}
			if (objectDefaults.lnDef !== undefined) {
				od['a:lnDef'] = objectDefaults.lnDef as XmlObject;
			}
			if (objectDefaults.txDef !== undefined) {
				od['a:txDef'] = objectDefaults.txDef as XmlObject;
			}
			themeRoot['a:objectDefaults'] = od;
		} else {
			// Empty objectDefaults is canonical per Office output.
			themeRoot['a:objectDefaults'] = {};
		}

		// extraClrSchemeLst.
		const extraClr = this.masterThemeExtraClrSchemeLst.get(themePath);
		themeRoot['a:extraClrSchemeLst'] =
			extraClr !== undefined ? (extraClr as XmlObject) : ({} as XmlObject);

		// custClrLst.
		const custClr = this.masterThemeCustClrLst.get(themePath);
		if (custClr !== undefined) {
			themeRoot['a:custClrLst'] = custClr as XmlObject;
		}

		// extLst.
		const themeExt = this.masterThemeExtLst.get(themePath);
		if (themeExt !== undefined) {
			themeRoot['a:extLst'] = themeExt as XmlObject;
		}

		const doc: XmlObject = {
			'?xml': { '@_version': '1.0', '@_encoding': 'UTF-8', '@_standalone': 'yes' },
			'a:theme': themeRoot,
		};

		try {
			return this.builder.build(doc);
		} catch (error) {
			console.warn(`Failed to build theme XML for ${themePath}:`, error);
			return undefined;
		}
	}

	/**
	 * Re-parse the original theme XML and pluck out a subtree by path,
	 * returning the raw parser object. Returns `undefined` when the
	 * original is missing or the path doesn't exist.
	 *
	 * Used to preserve `a:fmtScheme` byte-for-byte through a regenerate
	 * round-trip, since the in-memory PptxThemeFormatScheme is lossy.
	 */
	protected extractRawSubtreeFromOriginal(themePath: string, path: string[]): unknown | undefined {
		const original = this.originalThemeXmlByPath.get(themePath);
		if (!original) {
			return undefined;
		}
		try {
			const parsed = this.parser.parse(original) as XmlObject;
			let cursor: unknown = parsed;
			for (const segment of path) {
				if (cursor && typeof cursor === 'object' && segment in (cursor as XmlObject)) {
					cursor = (cursor as XmlObject)[segment];
				} else {
					return undefined;
				}
			}
			return cursor;
		} catch {
			return undefined;
		}
	}
}
