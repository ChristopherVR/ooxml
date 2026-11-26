import { copyParsedXmlWithOrder, deriveForLoad } from './load-part-reader';
import { XmlObject } from '../../types';
import type { PptxTheme, PptxThemeFormatScheme, PptxThemeOption } from '../../types';
import { PptxHandlerRuntime as PptxHandlerRuntimeBase } from './PptxHandlerRuntimeThemePartModel';
import { colorSchemeFromMap, fontSchemeFromMap } from './theme-part-model';
import { parseThemeColorMap } from '../../../../drawingml/theme';
import { NS, first, parseXml, type XmlElement } from '../../../../xml/index';
import { at } from '../../../../geometry/indexed';

interface ThemeXmlRead {
	parser: { parse(xml: string): unknown };
	source: string;
	tree: XmlObject;
	drawingRoot?: XmlElement | null;
	claimed: boolean;
}

export function isPresentationThemePartPath(path: string): boolean {
	return /^ppt\/theme\/(?!themeOverride)[^/]+\.xml$/iu.test(path);
}

/**
 * The `p:clrMap` of a slide master's text (ECMA-376 Part 1, 19.3.1.6) as the alias dictionary the
 * runtime keeps (`bg1` -> `lt1`...), read by the shared `parseThemeColorMap`: every value is an
 * `ST_ColorSchemeIndex` token in its spec spelling, a value outside the enumeration is skipped.
 * `undefined` when the master has no colour map (or the part does not parse).
 */
export function masterColorMapOf(masterXml: string): Record<string, string> | undefined {
	let root: XmlElement;
	try {
		root = parseXml(masterXml, { label: 'PPTX slide master' }).documentElement;
	} catch {
		return undefined;
	}
	if (root.namespaceURI !== NS.p || root.localName !== 'sldMaster') return undefined;
	const map: Record<string, string> = { ...parseThemeColorMap(first(root, 'clrMap', NS.p)) };
	return Object.keys(map).length > 0 ? map : undefined;
}

export class PptxHandlerRuntime extends PptxHandlerRuntimeBase {
	protected async resolvePrimaryThemePath(): Promise<string | undefined> {
		const masterFiles = this.zip.file(/^ppt\/slideMasters\/slideMaster\d+\.xml$/);
		if (!masterFiles || masterFiles.length === 0) {
			return undefined;
		}
		const masterPath = at(masterFiles, 0).name;
		const relsPath = masterPath.replace(
			/ppt\/slideMasters\/(slideMaster\d+)\.xml/,
			'ppt/slideMasters/_rels/$1.xml.rels',
		);
		const relsXml = this.zip.file(relsPath);
		if (!relsXml) {
			return undefined;
		}
		const relsData = await this.readPartXml(relsPath);
		const relNodes = this.ensureArray(
			(relsData?.Relationships as XmlObject | undefined)?.Relationship,
		) as XmlObject[];
		for (const rel of relNodes) {
			const target = String(rel['@_Target'] || '');
			if (!target.includes('theme')) {
				continue;
			}
			const themePath = target.startsWith('..')
				? this.resolvePath(masterPath.substring(0, masterPath.lastIndexOf('/') + 1), target)
				: target.startsWith('/')
					? target.slice(1)
					: `ppt/${target.replace(/^\.?\//, '')}`;
			if (themePath.startsWith('ppt/theme/')) {
				return themePath;
			}
		}
		return undefined;
	}

	protected async parseThemeOptions(): Promise<PptxThemeOption[]> {
		const themeFiles = Object.values(this.zip.files).filter(
			(file) => !file.dir && isPresentationThemePartPath(file.name),
		);
		if (!themeFiles || themeFiles.length === 0) {
			return [];
		}
		const options: PptxThemeOption[] = [];
		for (const file of themeFiles) {
			try {
				const xml = (await this.readPartText(file.name)) ?? '';
				const data = (await this.readThemeObject(file.name, xml)).tree;
				const root = data['a:theme'] as XmlObject | undefined;
				const name = String(root?.['@_name'] || '').trim();
				options.push({
					path: file.name,
					name: name.length > 0 ? name : undefined,
				});
			} catch {
				options.push({ path: file.name });
			}
		}
		return options;
	}

	/**
	 * Public API — scan the in-memory ZIP for `ppt/theme/theme*.xml` parts
	 * and return their paths and display names.  Delegates to the
	 * protected {@link parseThemeOptions}.
	 */
	public async getAvailableThemes(): Promise<PptxThemeOption[]> {
		return this.parseThemeOptions();
	}

	/**
	 * Build a structured PptxTheme object from the already-parsed
	 * themeColorMap and themeFontMap for consumption by renderers / UI.
	 */
	protected buildThemeObject(): PptxTheme | undefined {
		const hasColors = Object.keys(this.themeColorMap).length > 0;
		const hasFonts = Object.keys(this.themeFontMap).length > 0;
		if (!hasColors && !hasFonts) {
			return undefined;
		}

		// #83: surface the parsed per-script overrides (`<a:font script=...>`) on the theme object so
		// font resolution can consult them. The maps are keyed by theme path; aggregate across every
		// parsed theme (a single-master deck has one entry) into one script -> face lookup.
		const colorScheme = hasColors ? colorSchemeFromMap(this.themeColorMap) : undefined;
		const fontScheme = hasFonts
			? fontSchemeFromMap(
					this.themeFontMap,
					this.aggregateFontScriptOverrides(this.masterThemeMajorFontScripts),
					this.aggregateFontScriptOverrides(this.masterThemeMinorFontScripts),
				)
			: undefined;

		return {
			colorScheme,
			fontScheme,
			formatScheme: this.themeFormatScheme,
		};
	}

	/**
	 * Parse every slide master's `<p:clrMap>` element and store the alias
	 * dictionaries on {@link masterClrMaps}. Do *not* mutate
	 * {@link themeColorMap} — alias resolution happens at colour-lookup
	 * time so that:
	 *
	 * 1. The raw theme scheme stays the source of truth (clrMap is a
	 *    routing layer, not a colour table).
	 * 2. Multi-master decks resolve each slide against its own master's
	 *    clrMap rather than always against `masterFiles[0]`.
	 * 3. Layout `clrMapOvr` semantics work correctly when a slide's master
	 *    differs from the deck's first master.
	 *
	 * Phase 2 Stream B / C-H4.
	 */
	protected async applySlideMasterColorMap(_defaultMap: Record<string, string>): Promise<void> {
		const masterFiles = this.zip.file(/^ppt\/slideMasters\/slideMaster\d+\.xml$/);
		if (!masterFiles || masterFiles.length === 0) {
			return;
		}

		for (const file of masterFiles) {
			try {
				// eslint-disable-next-line no-await-in-loop
				const masterXml = (await this.readPartText(file.name)) ?? '';
				const aliasMap = masterColorMapOf(masterXml);
				if (aliasMap) {
					this.masterClrMaps.set(file.name, aliasMap);
				}
			} catch (error) {
				console.warn(`Failed to parse slide master color map for ${file.name}:`, error);
			}
		}
	}

	/**
	 * Parse a single theme part into structured colour, font, and format
	 * scheme dictionaries. Used both for the global default theme and for
	 * each master's per-master theme (multi-master support).
	 *
	 * Phase 2 Stream B / C-H4.
	 */
	protected async parseThemePart(themePath: string): Promise<{
		colorMap: Record<string, string>;
		fontMap: Record<string, string>;
		formatScheme: PptxThemeFormatScheme | undefined;
	} | null> {
		const themeFile = this.zip.file(themePath);
		if (!themeFile) {
			return null;
		}
		const themeXml = (await this.readPartText(themePath)) ?? '';
		const cached = await this.readThemeObject(themePath, themeXml);
		const tree = cached.claimed ? copyParsedXmlWithOrder(cached.tree, themeXml) : cached.tree;
		cached.claimed = true;
		if (cached.drawingRoot === undefined)
			cached.drawingRoot = this.readDrawingThemeRoot(themeXml) ?? null;
		const model = this.parseThemePartXml(themeXml, tree, cached.drawingRoot);
		this.cacheThemePartModel(themePath, themeXml, model);
		return { colorMap: model.colorMap, fontMap: model.fontMap, formatScheme: model.formatScheme };
	}

	/** Cache the XML reads, while each caller resolves styles in its own colour context. */
	private async readThemeObject(path: string, source: string) {
		const parser = this.parser;
		const parse = (): ThemeXmlRead => ({
			parser,
			source,
			tree: parser.parse(source) as XmlObject,
			claimed: false,
		});
		const cached = await deriveForLoad(this.zip, `theme-xml:${path}`, async () => parse());
		return cached.parser === parser && cached.source === source ? cached : parse();
	}

	/**
	 * Flatten a per-theme-path script-override map (`themePath -> {script ->
	 * typeface}`) into a single `{script -> typeface}` lookup for
	 * {@link buildThemeObject}. Earlier entries win on collision, which matches
	 * the primary-theme-first parse order. Phase 4 Stream A / M4 (#83).
	 */
	protected aggregateFontScriptOverrides(
		perPathMap: Map<string, Record<string, string>>,
	): Record<string, string> {
		const aggregate: Record<string, string> = {};
		for (const overrides of perPathMap.values()) {
			for (const [script, typeface] of Object.entries(overrides)) {
				if (!(script in aggregate)) {
					aggregate[script] = typeface;
				}
			}
		}
		return aggregate;
	}

	/**
	 * Resolve the theme file path referenced by a given master's `.rels`.
	 * Returns `undefined` when the master has no theme relationship.
	 */
	protected async resolveThemePathForMaster(masterPath: string): Promise<string | undefined> {
		const relsPath = masterPath.replace(
			/ppt\/slideMasters\/(slideMaster\d+)\.xml/,
			'ppt/slideMasters/_rels/$1.xml.rels',
		);
		const relsXml = this.zip.file(relsPath);
		if (!relsXml) {
			return undefined;
		}
		const relsData = await this.readPartXml(relsPath);
		const relNodes = this.ensureArray(
			(relsData?.Relationships as XmlObject | undefined)?.Relationship,
		) as XmlObject[];
		for (const rel of relNodes) {
			const target = String(rel['@_Target'] || '');
			if (!target.includes('theme')) {
				continue;
			}
			const themePath = target.startsWith('..')
				? this.resolvePath(masterPath.substring(0, masterPath.lastIndexOf('/') + 1), target)
				: target.startsWith('/')
					? target.slice(1)
					: `ppt/${target.replace(/^\.?\//, '')}`;
			if (themePath.startsWith('ppt/theme/')) {
				return themePath;
			}
		}
		return undefined;
	}

	/**
	 * Populate {@link masterThemeColorMaps}, {@link masterThemeFontMaps},
	 * and {@link masterThemeFormatSchemes} for every slide master in the
	 * deck. Multi-master support — Phase 2 Stream B / C-H4.
	 */
	protected async loadPerMasterThemes(): Promise<void> {
		const masterFiles = this.zip.file(/^ppt\/slideMasters\/slideMaster\d+\.xml$/);
		if (!masterFiles || masterFiles.length === 0) {
			return;
		}
		for (const file of masterFiles) {
			try {
				// eslint-disable-next-line no-await-in-loop
				const themePath = await this.resolveThemePathForMaster(file.name);
				if (!themePath) {
					continue;
				}
				this.masterThemePaths.set(file.name, themePath);
				// eslint-disable-next-line no-await-in-loop
				const parsed = await this.parseThemePart(themePath);
				if (!parsed) {
					continue;
				}
				this.masterThemeColorMaps.set(file.name, parsed.colorMap);
				this.masterThemeFontMaps.set(file.name, parsed.fontMap);
				if (parsed.formatScheme) {
					this.masterThemeFormatSchemes.set(file.name, parsed.formatScheme);
				}
			} catch (error) {
				console.warn(`Failed to load per-master theme for ${file.name}:`, error);
			}
		}
	}

	protected async loadThemeData(): Promise<void> {
		const themeFiles = this.zip.file(/^ppt\/theme\/theme\d+\.xml$/);
		if (!themeFiles || themeFiles.length === 0) {
			return;
		}

		const preferredThemePath = await this.resolvePrimaryThemePath();
		const themeFile = preferredThemePath
			? (themeFiles.find((file) => file.name === preferredThemePath) ?? at(themeFiles, 0))
			: at(themeFiles, 0);

		const parsed = await this.parseThemePart(themeFile.name);
		if (parsed) {
			this.themeColorMap = parsed.colorMap;
			this.themeFontMap = parsed.fontMap;
			if (parsed.formatScheme) {
				this.themeFormatScheme = parsed.formatScheme;
			}
		}
		// Parse master clrMaps and per-master themes. These layer on top of
		// {@link themeColorMap} during slide rendering rather than mutating it.
		await this.applySlideMasterColorMap(this.getDefaultSchemeColorMap());
		await this.loadPerMasterThemes();

		// Snapshot the post-load global theme so per-slide master switching
		// can restore it when a slide has no per-master entry. Phase 2
		// Stream B / C-H4.
		this.globalThemeColorMapSnapshot = { ...this.themeColorMap };
		this.globalThemeFontMapSnapshot = { ...this.themeFontMap };
		this.globalThemeFormatSchemeSnapshot = this.themeFormatScheme;
	}
}
