import { PptxHandlerRuntime as PptxHandlerRuntimeBase } from './PptxHandlerRuntimeThemeProcessing';

const SLIDE_MASTER_PART = /^ppt\/slideMasters\/slideMaster\d+\.xml$/;

/** `themePath` as a package path, or `undefined` when it does not name a theme part. */
function themePartPath(themePath: string): string | undefined {
	const path = themePath.trim().replace(/\\/g, '/');
	return path.startsWith('ppt/theme/') ? path : undefined;
}

/**
 * Pointing slide masters at another theme part already in the package (Presentation properties >
 * Theme).
 *
 * Rewriting a master's theme relationship is not enough on its own: the runtime's theme state
 * (`masterThemePaths`, the per-master colour, font and format maps, and for the first master the
 * active maps and the deck-wide snapshots) still described the old part, so a scheme edit
 * (`resolveThemeEditTargets`) and a save (`persistThemeParts`) made after the switch went to the
 * part the masters no longer use. Every re-targeted master adopts the new part's state too.
 */
export class PptxHandlerRuntime extends PptxHandlerRuntimeBase {
	/**
	 * Point the first slide master, or every one, at a theme part already in the package. Does
	 * nothing when the package has no such part.
	 */
	public async setPresentationTheme(themePath: string, applyToAllMasters = true): Promise<void> {
		const path = themePartPath(themePath);
		if (!path || !this.zip.file(path)) {
			return;
		}
		const masters = this.zip.file(SLIDE_MASTER_PART);
		for (const master of applyToAllMasters ? masters : masters.slice(0, 1)) {
			await this.retargetMasterTheme(master.name, path);
		}
	}

	/**
	 * Point one slide master at a theme part already in the package (Undo of a theme switch puts
	 * each master back on its own part). Does nothing when the package has no such master or part.
	 */
	public async setSlideMasterTheme(masterPath: string, themePath: string): Promise<void> {
		const path = themePartPath(themePath);
		if (!path || !this.zip.file(path) || !SLIDE_MASTER_PART.test(masterPath)) {
			return;
		}
		if (this.zip.file(masterPath)) {
			await this.retargetMasterTheme(masterPath, path);
		}
	}

	private async retargetMasterTheme(masterPath: string, themePath: string): Promise<void> {
		if (await this.setMasterThemeRelationship(masterPath, themePath)) {
			await this.adoptMasterTheme(masterPath, themePath);
		}
	}

	/** Point the runtime's theme state for `masterPath` at the part it now uses. */
	private async adoptMasterTheme(masterPath: string, themePath: string): Promise<void> {
		const themeXml = await this.zip.file(themePath)?.async('string');
		if (themeXml === undefined) {
			return;
		}
		const model = this.parseThemePartXml(themeXml);
		// A part no master used at load was never cached: the save needs its names and raw subtrees.
		if (!this.originalThemeXmlByPath.has(themePath)) {
			this.cacheThemePartModel(themePath, themeXml, model);
		}
		this.masterThemePaths.set(masterPath, themePath);
		this.masterThemeColorMaps.set(masterPath, { ...model.colorMap });
		this.masterThemeFontMaps.set(masterPath, { ...model.fontMap });
		if (model.formatScheme) {
			this.masterThemeFormatSchemes.set(masterPath, model.formatScheme);
		} else {
			this.masterThemeFormatSchemes.delete(masterPath);
		}
		// The first master's part is the deck's primary theme (`loadThemeData`).
		if (this.zip.file(SLIDE_MASTER_PART)[0]?.name !== masterPath) {
			return;
		}
		this.themeColorMap = { ...model.colorMap };
		this.themeFontMap = { ...model.fontMap };
		this.themeFormatScheme = model.formatScheme;
		this.globalThemeColorMapSnapshot = { ...model.colorMap };
		this.globalThemeFontMapSnapshot = { ...model.fontMap };
		this.globalThemeFormatSchemeSnapshot = model.formatScheme;
	}
}
