/**
 * Framework-neutral types for the ribbon style galleries (Shape Styles,
 * WordArt Styles, Picture Styles, Table Styles, Chart Styles, ...).
 *
 * A binding renders a gallery by asking `buildRibbonGallery(id, context)` for
 * a {@link RibbonGalleryDescriptor} and mapping it onto its own markup: a
 * trigger button, then a grid of tiles whose preview is the ready-made
 * `previewSvg` string. Clicking a tile calls
 * `applyRibbonGalleryItem(id, itemId, context)` and dispatches the returned
 * {@link RibbonGalleryApplyResult} to the binding's existing update path.
 * Everything that decides WHAT a gallery offers, how a tile looks and what
 * OOXML a pick writes lives in shared, so a new entry reaches all five
 * bindings at once.
 *
 * @module render/ribbon-galleries/gallery-types
 */
import type {
	PptxElement,
	PptxSlide,
	PptxTheme,
	PptxThemeVariant,
	PptxThemeColorScheme,
	PptxThemeFontScheme,
	ResolvedStyleMatrix,
	XmlObject,
} from 'ooxml-core/pptx';

/** Every gallery the ribbon offers. */
export type RibbonGalleryId =
	| 'shapeStyles'
	| 'shapeEffects'
	| 'wordArtStyles'
	/** The inspector's Text Warp (WordArt Transform) picker. */
	| 'textWarp'
	/** The inspector's Quick Styles (shape style presets) picker. */
	| 'quickStyles'
	| 'pictureStyles'
	| 'pictureCorrections'
	| 'pictureColor'
	| 'pictureArtisticEffects'
	| 'bullets'
	| 'numbering'
	| 'tableStyles'
	| 'tableSelect'
	| 'tableViewGridlines'
	| 'tableDeleteRows'
	| 'tableDeleteColumns'
	| 'tableInsertAbove'
	| 'tableInsertBelow'
	| 'tableInsertLeft'
	| 'tableInsertRight'
	| 'tableMergeCells'
	| 'tableSplitCells'
	| 'tableDistributeRows'
	| 'tableDistributeColumns'
	| 'tableAlignLeft'
	| 'tableAlignCenter'
	| 'tableAlignRight'
	| 'tableAlignTop'
	| 'tableAlignMiddle'
	| 'tableAlignBottom'
	| 'tableTextDirection'
	| 'tableCellMargins'
	| 'tableSize'
	| 'tableArrange'
	| 'chartStyles'
	| 'chartColors'
	| 'chartQuickLayout'
	| 'smartArtStyles'
	| 'smartArtColors'
	| 'smartArtLayouts'
	| 'smartArtAddShape'
	| 'smartArtAddShapeBefore'
	| 'smartArtAddShapeAbove'
	| 'smartArtAddBullet'
	| 'smartArtTextPane'
	| 'smartArtPromote'
	| 'smartArtDemote'
	| 'smartArtRightToLeft'
	| 'smartArtMoveUp'
	| 'smartArtMoveDown'
	| 'smartArtResetGraphic'
	| 'smartArtConvert'
	| 'themes'
	| 'themeVariants'
	| 'themeColors'
	| 'themeFonts'
	| 'themeEffects'
	| 'themeBackgrounds';

/** One tile. */
export interface RibbonGalleryItem {
	/** Stable within its gallery; what `applyRibbonGalleryItem` takes. */
	id: string;
	/** i18n key of the tile's accessible name / tooltip. */
	labelKey: string;
	/** `{{name}}` interpolation values for `labelKey`. */
	labelParams?: Readonly<Record<string, string | number>>;
	/**
	 * Interpolation values that are themselves translatable names ("Accent 1" in "Colored Fill -
	 * Accent 1"): param name -> `{ key, fallback }`, resolved before `labelKey` is filled.
	 */
	labelParamKeys?: Readonly<Record<string, { readonly key: string; readonly fallback: string }>>;
	/** English fallback for bindings whose dictionary lacks `labelKey`. */
	label: string;
	/**
	 * A complete, self-contained `<svg>` element (width/height/viewBox set).
	 * Built only from catalogue data and theme colours, never from user text,
	 * so it is safe to inject as markup.
	 */
	previewSvg: string;
	/**
	 * CSS `filter` the tile's preview carries (the Picture Adjust galleries),
	 * for a host that draws the thumbnail over a real image instead of
	 * `previewSvg`. Omitted when the preset has no CSS-expressible preview.
	 */
	previewFilter?: string;
	/** True when this entry is what the selection currently carries. */
	applied: boolean;
	/**
	 * Context menu entries (right-click, Menu key, Shift+F10); a plain click picks the tile with
	 * no action. Background Styles offers "Apply to All Slides" and "Apply to Selected Slides".
	 */
	actions?: RibbonGalleryAction[];
}

/** A named entry of a tile's context menu or of the dropdown footer. */
export interface RibbonGalleryAction {
	id: string;
	labelKey: string;
	label: string;
}

/** A titled block of tiles (PowerPoint's "Theme Styles" / "Presets"). */
export interface RibbonGallerySection {
	id: string;
	/** i18n key of the section heading; omitted for an untitled section. */
	titleKey?: string;
	title?: string;
	/** Tiles per row in the dropped-down grid. */
	columns: number;
	/** Preview tile size in CSS px (the SVG's own width/height). */
	tileWidth: number;
	tileHeight: number;
	items: RibbonGalleryItem[];
}

export interface RibbonGalleryDescriptor {
	id: RibbonGalleryId;
	/** i18n key of the gallery trigger / group caption. */
	labelKey: string;
	label: string;
	sections: RibbonGallerySection[];
	/** Tiles an inline gallery shows in the ribbon itself (default 6). */
	inlineCount?: number;
	/**
	 * True when hovering a tile shows its result on the slides without committing it (the Design
	 * galleries). The element emits `gallery-preview` / `gallery-preview-end`.
	 */
	previewable?: boolean;
	/** Commands under the tiles of the dropdown ("Format Background..."); a click picks `id`. */
	footer?: RibbonGalleryAction[];
	/** True when the selection cannot take this gallery (trigger disabled). */
	disabled: boolean;
	/**
	 * Set when the entry is a one-button command (PowerPoint's Add Shape, Reset
	 * Graphic, ...) rather than a gallery: the element draws a plain command, a
	 * click picks its only item and no panel opens.
	 */
	command?: RibbonGalleryCommandView;
	/**
	 * Set when part of the gallery loads lazily and was not ready when the descriptor was built.
	 * Resolves with the rebuilt descriptor once it has landed (or `undefined` when it never will),
	 * so the element can swap it in and a binding that memoizes its descriptor need not know.
	 */
	ready?: Promise<RibbonGalleryDescriptor | undefined>;
}

/** How a command-style entry draws (see {@link RibbonGalleryDescriptor.command}). */
export interface RibbonGalleryCommandView {
	/** SVG path data on a 20px grid. */
	iconPath: string;
	/** Icon above the label (PowerPoint's large button) instead of beside it. */
	large?: boolean;
	/** i18n key and English fallback of the tooltip, e.g. why the command is unavailable. */
	hintKey?: string;
	hint?: string;
}

/**
 * What a gallery needs to know about the deck and the selection. Bindings
 * build it once per render from state they already hold.
 */
export interface RibbonGalleryContext {
	/** The primary selected element, or null. */
	element: PptxElement | null;
	/** `PptxData.themeColorMap`: scheme key -> hex. */
	themeColorMap?: Readonly<Record<string, string>>;
	/** `PptxData.theme`. */
	theme?: PptxTheme;
	/**
	 * The handler's `resolveStyleMatrixReferences`, bound. Shape Styles use it
	 * to resolve `<p:style>` exactly as the load path does; without it they
	 * fall back to an approximation from `themeColorMap`.
	 */
	resolveStyleMatrix?: (styleXml: XmlObject, effectSchemeId?: string) => ResolvedStyleMatrix;
	/**
	 * The table cells the selection covers (`row` and `column` are zero-based), for the Table
	 * Layout commands. Omitted or empty when the table frame itself is selected.
	 */
	selectedTableCells?: ReadonlyArray<{ row: number; column: number }>;
	/** The slide being edited (Background Styles marks the style it uses). */
	activeSlide?: PptxSlide | null;
	/** `PptxData.themeVariants`: real variants from the package, when it has any. */
	themeVariants?: readonly PptxThemeVariant[];
}

/** What a pick asks the binding to do. */
export type RibbonGalleryApplyResult =
	| {
			kind: 'element';
			elementId: string;
			/** Merge onto the element with the binding's normal update + history path. */
			patch: Partial<PptxElement>;
	  }
	| {
			/** A whole theme: colour and font scheme together (Design > Themes). */
			kind: 'theme';
			colorScheme: PptxThemeColorScheme;
			fontScheme: PptxThemeFontScheme;
			name: string;
	  }
	| {
			kind: 'themeColorScheme';
			colorScheme: PptxThemeColorScheme;
			name: string;
	  }
	| {
			kind: 'themeFontScheme';
			fontScheme: PptxThemeFontScheme;
			name: string;
	  }
	| {
			/** Design > Variants > Effects: a built-in scheme, or a variant's own `a:fmtScheme`. */
			kind: 'themeEffects';
			schemeId?: string;
			fmtSchemeXml?: string;
			name: string;
	  }
	| {
			/** Design > Variants > Background Styles: preset 1 to 12. */
			kind: 'backgroundStyle';
			styleIndex: number;
			scope: 'all' | 'selected';
	  }
	| {
			/** A command the host runs (Format Background... opens its background pane). */
			kind: 'command';
			command: 'formatBackground';
	  }
	| {
			/** Several results in order (a real theme variant: colours and fonts, effects, background). */
			kind: 'sequence';
			steps: RibbonGalleryApplyResult[];
	  };
