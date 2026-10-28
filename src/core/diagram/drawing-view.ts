// The laid-out view of a SmartArt shape that renderers and editing operations work on: a shape of
// the cached `dsp:drawing` with its colours resolved to `#RRGGBB` and its geometry in the host's
// pixel space, or a shape produced by a reflow after an edit. PowerPoint's
// `PptxSmartArtDrawingShape` extends it with its 3D, rich-text segment and custom-geometry slots;
// Word and Excel can use it as is. Extracted from `pptx/core/types/smart-art.ts` (3b2ecd9a6).

/** A gradient stop of a resolved fill. `position` is a percentage (0..100). */
export interface DiagramShapeGradientStop {
	color: string;
	position: number;
	opacity?: number;
}

/** A SmartArt shape with resolved colours, in pixels (EMU / `emuPerPx`). */
export interface DiagramShapeView {
	/** Shape id: the `dsp:sp/@modelId` presentation point, or `reflow-<kind>-<nodeId>`. */
	id: string;
	/** Preset geometry (`roundRect`, `ellipse`), or `custom`. */
	shapeType?: string;
	x: number;
	y: number;
	width: number;
	height: number;
	/** Rotation in degrees. */
	rotation?: number;
	/** Horizontal/vertical mirrors declared on the cached drawing transform. */
	flipHorizontal?: boolean;
	flipVertical?: boolean;
	/** Skew along the X axis in degrees. */
	skewX?: number;
	/** Skew along the Y axis in degrees. */
	skewY?: number;
	/** Preset-geometry adjustment values from `a:prstGeom/a:avLst`. */
	shapeAdjustments?: Record<string, number>;
	/**
	 * The cached shape declares `a:noFill`. Renderers must leave it unpainted rather than
	 * substituting a palette colour, because these shapes usually sit on top of a painted shape
	 * whose fill has to stay visible.
	 */
	fillNone?: boolean;
	/** Solid fill colour (hex). */
	fillColor?: string;
	/**
	 * Opacity (0..1) of the solid fill, from the colour's `a:alpha` (Basic Venn paints its circles
	 * at 50% so overlaps blend). Absent when opaque.
	 */
	fillOpacity?: number;
	/** Gradient fill stops when the cached shape uses `a:gradFill` (positions 0..100). */
	fillGradientStops?: DiagramShapeGradientStop[];
	/** Gradient geometry type (`linear` for `a:lin`, `radial` for `a:path`). */
	fillGradientType?: 'linear' | 'radial';
	/** Linear gradient angle in degrees (0..360). */
	fillGradientAngle?: number;
	/** Pattern fill preset name from `a:pattFill/@prst` (e.g. "pct50", "cross"). */
	fillPatternPreset?: string;
	/** Pattern fill foreground colour (hex) from `a:pattFill/a:fgClr`. */
	fillPatternForegroundColor?: string;
	/** Pattern fill background colour (hex) from `a:pattFill/a:bgClr`. */
	fillPatternBackgroundColor?: string;
	/** Relationship id of a picture (blip) fill's embedded image (`a:blip/@r:embed`). */
	fillBlipEmbedId?: string;
	/** Resolved data-URI/URL for a picture fill, when the host could resolve the image part. */
	fillImageUrl?: string;
	fillImageWidth?: number;
	fillImageHeight?: number;
	/** Complete source-backed image fill, including crop, effects and layout semantics. */
	pictureFill?: Extract<import('../drawingml/types').DrawingFill, { kind: 'picture' }>;
	/** Whether the cached shape carries an outer-shadow effect (`a:effectLst`). */
	hasShadow?: boolean;
	/** Resolved outer-shadow colour (hex), when present. */
	shadowColor?: string;
	/** Stroke colour (hex). */
	strokeColor?: string;
	/** Stroke width in points. */
	strokeWidth?: number;
	/** Text content of the shape (paragraphs joined by a newline). */
	text?: string;
	/** Font size in CSS pixels. */
	fontSize?: number;
	/** Font colour (hex). */
	fontColor?: string;
	/** Authored font family from the first styled run. */
	fontFamily?: string;
	/** Authored font weight (400 or 700). */
	fontWeight?: number;
	/** Authored font style. */
	fontStyle?: 'normal' | 'italic';
	/** Absolute line height in CSS pixels from `a:spcPts`. */
	lineHeight?: number;
	/** Relative line height from `a:spcPct`. */
	lineHeightRatio?: number;
	/** Absolute spacing after a paragraph in CSS pixels. */
	lineSpacingAfter?: number;
	/** Relative spacing after a paragraph. */
	lineSpacingAfterRatio?: number;
	/** Text-body insets in CSS pixels. */
	textInsetLeft?: number;
	textInsetTop?: number;
	textInsetRight?: number;
	textInsetBottom?: number;
	/** DiagramML text vertical anchor (`t`, `ctr`, or `b`). */
	textVerticalAnchor?: string;
	/** Independent DiagramML text-frame geometry from `dsp:txXfrm`. */
	textFrameX?: number;
	textFrameY?: number;
	textFrameWidth?: number;
	textFrameHeight?: number;
	/**
	 * The presentation style label (`presStyleLbl`: `node1`, `revTx`, `sibTrans2D1`, ...) of the
	 * layout node a regenerated shape was laid out from, when the layout engine reports it;
	 * `undefined` on cached (parsed) shapes.
	 */
	styleLabel?: string;
}
