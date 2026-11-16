// The laid-out view of a SmartArt shape that renderers and editing operations work on: a shape of
// the cached `dsp:drawing` with its colours resolved to `#RRGGBB` and its geometry in the host's
// pixel space, or a shape produced by a reflow after an edit. PowerPoint's
// `PptxSmartArtDrawingShape` extends it with its 3D, rich-text segment and custom-geometry slots;
// Word and Excel can use it as is. Extracted from `pptx/core/types/smart-art.ts` (3b2ecd9a6).

/** A gradient stop of a resolved fill. `position` is a percentage (0..100). */
export interface DiagramShapeGradientStop {
	color: string;
	position: number;
	opacity?: number | undefined;
}

/** A SmartArt shape with resolved colours, in pixels (EMU / `emuPerPx`). */
export interface DiagramShapeView {
	/** Shape id: the `dsp:sp/@modelId` presentation point, or `reflow-<kind>-<nodeId>`. */
	id: string;
	/** Preset geometry (`roundRect`, `ellipse`), or `custom`. */
	shapeType?: string | undefined;
	x: number;
	y: number;
	width: number;
	height: number;
	/** Rotation in degrees. */
	rotation?: number | undefined;
	/** Horizontal/vertical mirrors declared on the cached drawing transform. */
	flipHorizontal?: boolean | undefined;
	flipVertical?: boolean | undefined;
	/** Skew along the X axis in degrees. */
	skewX?: number | undefined;
	/** Skew along the Y axis in degrees. */
	skewY?: number | undefined;
	/** Preset-geometry adjustment values from `a:prstGeom/a:avLst`. */
	shapeAdjustments?: Record<string, number> | undefined;
	/**
	 * The cached shape declares `a:noFill`. Renderers must leave it unpainted rather than
	 * substituting a palette colour, because these shapes usually sit on top of a painted shape
	 * whose fill has to stay visible.
	 */
	fillNone?: boolean | undefined;
	/** Solid fill colour (hex). */
	fillColor?: string | undefined;
	/**
	 * Opacity (0..1) of the solid fill, from the colour's `a:alpha` (Basic Venn paints its circles
	 * at 50% so overlaps blend). Absent when opaque.
	 */
	fillOpacity?: number | undefined;
	/** Gradient fill stops when the cached shape uses `a:gradFill` (positions 0..100). */
	fillGradientStops?: DiagramShapeGradientStop[] | undefined;
	/** Gradient geometry type (`linear` for `a:lin`, `radial` for `a:path`). */
	fillGradientType?: 'linear' | 'radial' | undefined;
	/** Linear gradient angle in degrees (0..360). */
	fillGradientAngle?: number | undefined;
	/** Pattern fill preset name from `a:pattFill/@prst` (e.g. "pct50", "cross"). */
	fillPatternPreset?: string | undefined;
	/** Pattern fill foreground colour (hex) from `a:pattFill/a:fgClr`. */
	fillPatternForegroundColor?: string | undefined;
	/** Pattern fill background colour (hex) from `a:pattFill/a:bgClr`. */
	fillPatternBackgroundColor?: string | undefined;
	/** Relationship id of a picture (blip) fill's embedded image (`a:blip/@r:embed`). */
	fillBlipEmbedId?: string | undefined;
	/** Resolved data-URI/URL for a picture fill, when the host could resolve the image part. */
	fillImageUrl?: string | undefined;
	fillImageWidth?: number | undefined;
	fillImageHeight?: number | undefined;
	/** Complete source-backed image fill, including crop, effects and layout semantics. */
	pictureFill?: Extract<import('../drawingml/types').DrawingFill, { kind: 'picture' }> | undefined;
	/** Whether the cached shape carries an outer-shadow effect (`a:effectLst`). */
	hasShadow?: boolean | undefined;
	/** Resolved outer-shadow colour (hex), when present. */
	shadowColor?: string | undefined;
	/** Stroke colour (hex). */
	strokeColor?: string | undefined;
	/** Stroke width in points. */
	strokeWidth?: number | undefined;
	/** Text content of the shape (paragraphs joined by a newline). */
	text?: string | undefined;
	/** Font size in CSS pixels. */
	fontSize?: number | undefined;
	/** Font colour (hex). */
	fontColor?: string | undefined;
	/** Authored font family from the first styled run. */
	fontFamily?: string | undefined;
	/** Authored font weight (400 or 700). */
	fontWeight?: number | undefined;
	/** Authored font style. */
	fontStyle?: 'normal' | 'italic' | undefined;
	/** Absolute line height in CSS pixels from `a:spcPts`. */
	lineHeight?: number | undefined;
	/** Relative line height from `a:spcPct`. */
	lineHeightRatio?: number | undefined;
	/** Absolute spacing after a paragraph in CSS pixels. */
	lineSpacingAfter?: number | undefined;
	/** Relative spacing after a paragraph. */
	lineSpacingAfterRatio?: number | undefined;
	/** Text-body insets in CSS pixels. */
	textInsetLeft?: number | undefined;
	textInsetTop?: number | undefined;
	textInsetRight?: number | undefined;
	textInsetBottom?: number | undefined;
	/** DiagramML text vertical anchor (`t`, `ctr`, or `b`). */
	textVerticalAnchor?: string | undefined;
	/** Independent DiagramML text-frame geometry from `dsp:txXfrm`. */
	textFrameX?: number | undefined;
	textFrameY?: number | undefined;
	textFrameWidth?: number | undefined;
	textFrameHeight?: number | undefined;
	/**
	 * The presentation style label (`presStyleLbl`: `node1`, `revTx`, `sibTrans2D1`, ...) of the
	 * layout node a regenerated shape was laid out from, when the layout engine reports it;
	 * `undefined` on cached (parsed) shapes.
	 */
	styleLabel?: string | undefined;
}
