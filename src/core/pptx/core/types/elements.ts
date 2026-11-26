/**
 * Concrete element types (one per `type` discriminant) and the
 * {@link PptxElement} discriminated union.
 *
 * Narrow on `element.type` to access variant-specific properties:
 * ```ts
 * if (element.type === "image") {
 *   console.log(element.imageData); // ImagePptxElement-only
 * }
 * ```
 *
 * @module pptx-types/elements
 */

// ==========================================================================
// Concrete element types (discriminated by `type`) and union
// ==========================================================================

import type { PptxChartData } from './chart';
import type { XmlObject } from './common';
import type { PptxElementBase, PptxTextProperties, PptxShapeProperties } from './element-base';
import type { PptxCustomPathProperties } from './geometry';
import type { PptxImageEffects, PptxImageProperties } from './image';
import type {
	PptxMediaType,
	PptxMediaReferenceKind,
	PptxAudioCdPosition,
	MediaBookmark,
	MediaMetadata,
	MediaCaptionTrack,
} from './media';
import type { Model3DSceneData } from './model3d-scene';
import type { ShapeStyle } from './shape-style';
import type { PptxSmartArtData } from './smart-art';
import type { PptxTableData } from './table';

/**
 * Accessibility metadata from `p:cNvPr/a:extLst`'s "Mark as decorative"
 * vendor extension (issue G16). PowerPoint's Alt Text pane writes
 * `a:ext[@uri='{C183D7F6-B498-43B3-948B-1728B52AA6E4}']/adec:decorative
 * val="1"` when a shape or image is marked decorative; mixed into the
 * element variants whose `p:cNvPr` that pane covers.
 */
export interface PptxAccessibilityProperties {
	/**
	 * Whether the element is marked decorative. When true, alt text /
	 * aria-label / Markdown export should skip describing the element even
	 * when {@link PptxImageProperties.altText} (or similar) is present.
	 */
	isDecorative?: boolean | undefined;
}

/**
 * Accessibility description/title from `p:cNvPr/@descr` / `@title` on a
 * plain shape, text box or connector (`p:sp` / `p:cxnSp`). The same pair of
 * attributes already round-trips for a graphic frame (see
 * {@link TablePptxElement.altText}) and, `descr` only, for a picture
 * ({@link PptxImageProperties.altText}); this mixin extends it to the three
 * element kinds whose PowerPoint Alt Text pane data was previously dropped
 * on load because neither field existed on the model.
 */
export interface PptxNonVisualDescription {
	/** `p:cNvPr/@descr`. */
	altText?: string | undefined;
	/** `p:cNvPr/@title`. */
	title?: string | undefined;
}

/**
 * `a:cNvPicPr/@preferRelativeResize` (issue G13), a picture-only non-visual
 * property distinct from `a:picLocks`.
 */
export interface PptxPictureNonVisualProperties {
	/** Office live-feed picture. Camera state is local to the viewer. */
	liveFeed?: boolean | undefined;
	/**
	 * `a:cNvPicPr/@preferRelativeResize` (ST_Boolean, defaults to `true` when
	 * absent). Controls whether a picture's crop rectangle is reinterpreted
	 * relative to the picture's ORIGINAL dimensions or its CURRENT
	 * (already-resized) dimensions when it is resized again after being
	 * cropped. Parsed and round-tripped for now; not yet wired into
	 * resize-after-crop arithmetic (this app always uses current-size
	 * semantics, which only diverges from `preferRelativeResize="0"` on a
	 * second resize after a crop).
	 */
	preferRelativeResize?: boolean | undefined;
}

/**
 * A text box — a plain rectangle containing text, typically with no
 * visible fill or stroke.
 *
 * @example
 * ```ts
 * const title: TextPptxElement = {
 *   type: "text",
 *   id: "txt_1", x: 50, y: 30, width: 800, height: 60,
 *   text: "Welcome",
 *   textStyle: { fontSize: 36, bold: true },
 * };
 * // => satisfies TextPptxElement
 * ```
 */
export interface TextPptxElement
	extends PptxElementBase, PptxTextProperties, PptxShapeProperties, PptxNonVisualDescription {
	type: 'text';
}

/**
 * A shape — may contain text and custom geometry (preset or freeform).
 *
 * @example
 * ```ts
 * const rect: ShapePptxElement = {
 *   type: "shape",
 *   id: "shp_1", x: 100, y: 200, width: 300, height: 150,
 *   shapeType: "roundRect",
 *   shapeStyle: { fillColor: "#00AA55" },
 *   text: "OK",
 * };
 * // => satisfies ShapePptxElement
 * ```
 */
export interface ShapePptxElement
	extends
		PptxElementBase,
		PptxTextProperties,
		PptxShapeProperties,
		PptxCustomPathProperties,
		PptxAccessibilityProperties,
		PptxNonVisualDescription {
	type: 'shape';
}

/**
 * A connector (straight, bent, or curved line between shapes).
 *
 * Connector endpoints can snap to specific shapes via
 * `shapeStyle.connectorStartConnection` / `connectorEndConnection`.
 *
 * @example
 * ```ts
 * const line: ConnectorPptxElement = {
 *   type: "connector",
 *   id: "cxn_1", x: 100, y: 100, width: 200, height: 0,
 *   shapeStyle: {
 *     strokeColor: "#333",
 *     connectorEndArrow: "triangle",
 *   },
 * };
 * // => satisfies ConnectorPptxElement
 * ```
 */
export interface ConnectorPptxElement
	extends PptxElementBase, PptxTextProperties, PptxShapeProperties, PptxNonVisualDescription {
	type: 'connector';
}

/**
 * An image element from an OOXML `<p:pic>` node with `type: "image"`.
 *
 * @example
 * ```ts
 * const img: ImagePptxElement = {
 *   type: "image",
 *   id: "img_1", x: 0, y: 0, width: 960, height: 540,
 *   imagePath: "ppt/media/image1.png",
 *   altText: "Background scenery",
 * };
 * // => satisfies ImagePptxElement
 * ```
 */
export interface ImagePptxElement
	extends
		PptxElementBase,
		PptxShapeProperties,
		PptxCustomPathProperties,
		PptxImageProperties,
		PptxAccessibilityProperties,
		PptxPictureNonVisualProperties {
	type: 'image';
}

/**
 * A picture element from an OOXML `<p:pic>` node with `type: "picture"`.
 *
 * Functionally identical to {@link ImagePptxElement} but distinguished by
 * the `type` discriminant for semantic clarity.
 */
export interface PicturePptxElement
	extends
		PptxElementBase,
		PptxShapeProperties,
		PptxCustomPathProperties,
		PptxImageProperties,
		PptxAccessibilityProperties,
		PptxPictureNonVisualProperties {
	type: 'picture';
}

/**
 * A single unrecognised `<a:graphicData>/<a:extLst>/<a:ext>` extension on a
 * graphicFrame, captured verbatim so the round-trip can preserve future or
 * vendor-specific markup that the parser doesn't yet understand.
 *
 * The XML is preserved as a fast-xml-parser object tree (the same shape as
 * `rawXml` on other elements) so the save layer can re-emit it through the
 * existing builder without lossy string manipulation.
 */
export interface PptxGraphicFrameExtension {
	/** The `@_uri` attribute identifying the extension (e.g. `{C3CD43...}`). */
	uri: string;
	/** Parsed XML payload of the extension, suitable for re-serialization. */
	xml: import('./common').XmlObject;
}

/**
 * A table embedded via a `<p:graphicFrame>`.
 *
 * @example
 * ```ts
 * const tbl: TablePptxElement = {
 *   type: "table",
 *   id: "tbl_1", x: 50, y: 200, width: 860, height: 300,
 *   tableData: {
 *     rows: [
 *       { cells: [{ text: "Name" }, { text: "Score" }] },
 *       { cells: [{ text: "Alice" }, { text: "95" }] },
 *     ],
 *   },
 * };
 * // => satisfies TablePptxElement
 * ```
 */
export interface TablePptxElement extends PptxElementBase {
	type: 'table';
	/** Parsed table cell data for editing. */
	tableData?: PptxTableData | undefined;
	/**
	 * Accessibility description from `p:nvGraphicFramePr/p:cNvPr/@descr`, the
	 * same non-visual-properties attribute a picture's alt text comes from.
	 */
	altText?: string | undefined;
	/** Accessibility title from `p:nvGraphicFramePr/p:cNvPr/@title`. */
	title?: string | undefined;
	/**
	 * Unrecognised extensions captured from `a:graphicData/a:extLst` so they
	 * round-trip losslessly. See {@link PptxGraphicFrameExtension}.
	 */
	extensionXml?: PptxGraphicFrameExtension[] | undefined;
}

/**
 * A chart embedded via a `<p:graphicFrame>`.
 *
 * Chart data is parsed from the related `chartN.xml` / `chartExN.xml`
 * parts inside the PPTX archive.
 */
export interface ChartPptxElement extends PptxElementBase {
	type: 'chart';
	chartData?: PptxChartData | undefined;
	/** Accessibility description from `p:nvGraphicFramePr/p:cNvPr/@descr`. */
	altText?: string | undefined;
	/** Accessibility title from `p:nvGraphicFramePr/p:cNvPr/@title`. */
	title?: string | undefined;
	/** Unrecognised graphicFrame extLst extensions, captured verbatim for round-trip. */
	extensionXml?: PptxGraphicFrameExtension[] | undefined;
}

/**
 * A SmartArt diagram embedded via a `<p:graphicFrame>`.
 *
 * SmartArt data is extracted from `dgm:dataModel` parts. The editor
 * supports real structural editing (adding, removing and reordering nodes,
 * editing node text, and switching layout presets) with a lossless
 * `ptLst` round-trip. When the file carries PowerPoint's own pre-computed
 * drawing part, that exact layout is used; otherwise an algorithmic layout
 * engine approximates it, so complex custom layouts may not match
 * PowerPoint pixel-for-pixel.
 */
export interface SmartArtPptxElement extends PptxElementBase {
	type: 'smartArt';
	smartArtData?: PptxSmartArtData | undefined;
	/** Accessibility description from `p:nvGraphicFramePr/p:cNvPr/@descr`. */
	altText?: string | undefined;
	/** Accessibility title from `p:nvGraphicFramePr/p:cNvPr/@title`. */
	title?: string | undefined;
	/** Unrecognised graphicFrame extLst extensions, captured verbatim for round-trip. */
	extensionXml?: PptxGraphicFrameExtension[] | undefined;
}

/**
 * Recognised OLE object application types derived from `progId` / `clsId`.
 *
 * Used to show type-specific icons and previews in the editor.
 */
export type OleObjectType =
	| 'excel'
	| 'word'
	| 'powerpoint'
	| 'pdf'
	| 'visio'
	| 'mathtype'
	| 'package'
	| 'unknown';

/**
 * An OLE (Object Linking and Embedding) object.
 *
 * OLE objects can be embedded Excel sheets, Word documents, PDFs, Visio
 * diagrams, MathType equations, or generic "packages". They carry a
 * preview image for display and optional binary data for extraction.
 *
 * @example
 * ```ts
 * const ole: OlePptxElement = {
 *   type: "ole",
 *   id: "ole_1", x: 100, y: 200, width: 400, height: 300,
 *   oleObjectType: "excel",
 *   oleProgId: "Excel.Sheet.12",
 *   fileName: "budget.xlsx",
 * };
 * // => satisfies OlePptxElement
 * ```
 */
export interface OlePptxElement extends PptxElementBase {
	type: 'ole';
	oleTarget?: string | undefined;
	oleProgId?: string | undefined;
	oleName?: string | undefined;
	/** CLSID of the OLE object (from `@_classid`). */
	oleClsId?: string | undefined;
	/** Detected application type (excel, word, pdf, etc.). */
	oleObjectType?: OleObjectType | undefined;
	/** File extension for the embedded binary (e.g. "xlsx", "docx"). */
	oleFileExtension?: string | undefined;
	/** Original file name when available. */
	fileName?: string | undefined;
	/** Whether this is a linked (vs. embedded) object. */
	isLinked?: boolean | undefined;
	/** External file path for linked OLE objects (TargetMode="External"). */
	externalPath?: string | undefined;
	/** Data-URL or path for the OLE preview image. */
	previewImage?: string | undefined;
	/**
	 * Archive path of the preview picture, resolved against the slide on load. With
	 * `eagerDecodeImages: false` (the default) the preview is not decoded during load; a host
	 * resolves `previewImageData` from this path, as it does a picture's `imagePath`.
	 */
	previewImagePath?: string | undefined;
	/** Decoded preview image as a data-URL. */
	previewImageData?: string | undefined;
	/** Whether the OLE object is shown as an icon (`p:oleObj/@showAsIcon`). */
	oleShowAsIcon?: boolean | undefined;
	/** Authored display width of the OLE object preview, in EMU (`@imgW`). */
	oleImgW?: number | undefined;
	/** Authored display height of the OLE object preview, in EMU (`@imgH`). */
	oleImgH?: number | undefined;
	/**
	 * The recovered embedded payload as a data-URL (e.g.
	 * `data:application/vnd...;base64,...`), suitable for download or
	 * open-in-new-tab. For a generic "Package" OLE object this is the unwrapped
	 * inner file; for a plain embedded file (e.g. `.xlsx`) it is that file
	 * directly. Undefined when the embedding is missing or unreadable.
	 *
	 * Stored as a data-URL string to mirror how images store decoded bytes
	 * ({@link ImagePptxElement.imageData}) and to stay serialization-safe.
	 */
	oleEmbeddedData?: string | undefined;
	/** Original file name of the embedded payload when recoverable. */
	oleEmbeddedFileName?: string | undefined;
	/** MIME type of the embedded payload, derived from its extension/ProgID. */
	oleEmbeddedMimeType?: string | undefined;
	/** Size of the embedded payload in bytes. */
	oleEmbeddedByteSize?: number | undefined;
	/**
	 * `p:link/@followColorScheme` (`ST_OleObjectFollowColorScheme`): whether a
	 * LINKED OLE object's icon recolours to match the presentation theme.
	 * Only meaningful when {@link isLinked} is `true`. ECMA-376 §19.3.1.28.
	 */
	oleFollowColorScheme?: 'none' | 'full' | 'textAndBackground' | undefined;
	/**
	 * `p:link/@updateAutomatic` (`CT_OleObjectLink`, ECMA-376 §19.3.2.4):
	 * whether a LINKED OLE object refreshes automatically from its source
	 * (PowerPoint's Edit Links dialog "Automatic" vs. "Manual" radio buttons).
	 * Only meaningful when {@link isLinked} is `true`. The schema default is
	 * `false`; `undefined` means the source authored no explicit value.
	 */
	oleUpdateAutomatic?: boolean | undefined;
	/**
	 * Set by the in-viewer OLE content editors (`ole-edit-api.ts`) whenever
	 * `oleEmbeddedData` and/or `previewImageData` have been changed in memory
	 * and still need to be written back into the saved package. Never
	 * authored from a parsed file; purely an in-memory save signal, mirroring
	 * `PptxSlide.isDirty`. The save writer clears it once the pending write
	 * has been queued.
	 */
	oleContentDirty?: boolean | undefined;
	/** Accessibility description from `p:nvGraphicFramePr/p:cNvPr/@descr`. */
	altText?: string | undefined;
	/** Accessibility title from `p:nvGraphicFramePr/p:cNvPr/@title`. */
	title?: string | undefined;
	/** Unrecognised graphicFrame extLst extensions, captured verbatim for round-trip. */
	extensionXml?: PptxGraphicFrameExtension[] | undefined;
}

/**
 * An audio or video media element.
 *
 * Media elements reference files inside the PPTX archive
 * (`mediaPath`) and may include trim points, poster frames, and
 * playback settings for presentation mode.
 *
 * @example
 * ```ts
 * const video: MediaPptxElement = {
 *   type: "media",
 *   id: "vid_1", x: 50, y: 100, width: 640, height: 360,
 *   mediaType: "video",
 *   mediaPath: "ppt/media/media1.mp4",
 *   autoPlay: true,
 *   volume: 0.8,
 * };
 * // => satisfies MediaPptxElement
 * ```
 */
export interface MediaPptxElement extends PptxElementBase {
	type: 'media';
	mediaType?: PptxMediaType | undefined;
	mediaPath?: string | undefined;
	mediaData?: string | undefined;
	mediaMimeType?: string | undefined;
	mediaReferenceKind?: PptxMediaReferenceKind | undefined;
	mediaReferenceName?: string | undefined;
	/** Explicit DrawingML `audioFile/@contentType` value when present. */
	mediaReferenceContentType?: string | undefined;
	audioCdStart?: PptxAudioCdPosition | undefined;
	audioCdEnd?: PptxAudioCdPosition | undefined;
	rawMediaReferenceXml?: XmlObject | undefined;
	/** Trim start in milliseconds from the start of the clip. */
	trimStartMs?: number | undefined;
	/** Trim end in milliseconds measured from the end of the clip. */
	trimEndMs?: number | undefined;
	/** Path to the poster/preview image inside the ZIP. */
	posterFramePath?: string | undefined;
	/** Base64 data-URL for the poster frame image. */
	posterFrameData?: string | undefined;
	/** DrawingML treatment of the media picture. */
	imageEffects?: PptxImageEffects | undefined;
	shapeStyle?: ShapeStyle | undefined;
	shapeType?: string | undefined;
	shapeAdjustments?: Record<string, number> | undefined;
	/** Poster source crop from the left edge as a 0..1 fraction. */
	cropLeft?: number | undefined;
	/** Poster source crop from the top edge as a 0..1 fraction. */
	cropTop?: number | undefined;
	/** Poster source crop from the right edge as a 0..1 fraction. */
	cropRight?: number | undefined;
	/** Poster source crop from the bottom edge as a 0..1 fraction. */
	cropBottom?: number | undefined;
	/** Poster stretch-target inset from the left frame edge. */
	fillRectLeft?: number | undefined;
	/** Poster stretch-target inset from the top frame edge. */
	fillRectTop?: number | undefined;
	/** Poster stretch-target inset from the right frame edge. */
	fillRectRight?: number | undefined;
	/** Poster stretch-target inset from the bottom frame edge. */
	fillRectBottom?: number | undefined;
	/** Whether media should play full-screen during presentation. */
	fullScreen?: boolean | undefined;
	/** Return to the initial frame when playback finishes. */
	rewindAfterPlaying?: boolean | undefined;
	/** Whether media should loop continuously. */
	loop?: boolean | undefined;
	/** Fade-in duration in seconds. */
	fadeInDuration?: number | undefined;
	/** Fade-out duration in seconds. */
	fadeOutDuration?: number | undefined;
	/** Playback volume (0 to 1). */
	volume?: number | undefined;
	/** The authored mute flag, independent of the clip's saved volume. */
	muted?: boolean | undefined;
	/** Whether media auto-plays on slide entry. */
	autoPlay?: boolean | undefined;
	/** Where a clip starts in the slide-show timeline. */
	startMode?: 'auto' | 'sequence' | 'click' | undefined;
	/** Whether audio continues playing across slide transitions (presentation mode). */
	playAcrossSlides?: boolean | undefined;
	/** Hide the element when media is not actively playing. */
	hideWhenNotPlaying?: boolean | undefined;
	/** Named time bookmarks within the clip. */
	bookmarks?: MediaBookmark[] | undefined;
	/** Playback speed multiplier (1 = normal, 2 = double, 0.5 = half). */
	playbackSpeed?: number | undefined;
	/** Runtime-extracted metadata (duration, resolution, codec). */
	metadata?: MediaMetadata | undefined;
	/** Closed caption / subtitle tracks. */
	captionTracks?: MediaCaptionTrack[] | undefined;
	/** Whether the media source is missing/broken (file not found in archive). */
	mediaMissing?: boolean | undefined;
	/**
	 * Whether the media is linked (external `r:link`) rather than embedded
	 * (`r:embed`). Defaults to embedded when undefined.
	 */
	isLinked?: boolean | undefined;
	/**
	 * Accessibility description. Read from `p:nvGraphicFramePr/p:cNvPr/@descr`
	 * for the `p:graphicFrame`-shaped (SDK-created) media form, or from
	 * `p:nvPicPr/p:cNvPr/@descr` for the `p:pic`-shaped media form (real
	 * PowerPoint's usual authoring shape for a video/audio placeholder); see
	 * `PptxHandlerRuntimePictureParsing.ts`.
	 */
	altText?: string | undefined;
	/** Accessibility title, from the same `@title` attribute on whichever `p:cNvPr` the media form uses. Same scope note as {@link altText}. */
	title?: string | undefined;
	/** Unrecognised graphicFrame extLst extensions, captured verbatim for round-trip. */
	extensionXml?: PptxGraphicFrameExtension[] | undefined;
}

/**
 * A group container that holds child elements.
 *
 * Children inherit the group’s transform, so moving/resizing the group
 * affects all children proportionally.
 *
 * @example
 * ```ts
 * const group: GroupPptxElement = {
 *   type: "group",
 *   id: "grp_1", x: 0, y: 0, width: 960, height: 540,
 *   children: [textEl, shapeEl],
 * };
 * // => satisfies GroupPptxElement
 * ```
 */
export interface GroupPptxElement extends PptxElementBase {
	type: 'group';
	/** Child elements contained within this group. */
	children: PptxElement[];
	/** Fill style extracted from the group's `p:grpSpPr`, used for `a:grpFill` inheritance. */
	groupFill?: ShapeStyle | undefined;
	/**
	 * The SAME `p:grpSpPr` extraction as {@link groupFill}, kept whenever the
	 * group carries a `p:grpSpPr` at all, regardless of whether it resolved to
	 * a paintable fill.
	 *
	 * `groupFill` is `undefined` unless the group has a real fill, because
	 * `getGroupChildParentFill`/`groupChildInheritedFill` (the `a:grpFill`
	 * inheritance chain) must keep chaining through an ancestor's fill when
	 * THIS group has none of its own. A group whose `p:grpSpPr` authors only
	 * `a:effectLst` (shadow/glow/soft-edge/reflection, no fill) needs those
	 * effects to still reach the renderer, so they are kept here under a name
	 * that carries no fill-inheritance meaning. Currently only reflection is
	 * read from it (`getComputedEffectStyle`); the rest of `a:effectLst` on a
	 * group remains unsupported.
	 */
	groupEffectStyle?: ShapeStyle | undefined;
	/**
	 * Exact EMU the group's own `a:chOff`/`a:chExt` (the coordinate space its
	 * CHILDREN are authored in) were parsed from, alongside {@link
	 * PptxElementBase.xEmu} etc for the group's own placement in its PARENT's
	 * space. `undefined` when the source carried no usable `a:chOff`/`a:chExt`
	 * (an SDK-created group, or one whose `a:xfrm` had no child-space data).
	 *
	 * Used by `group-xfrm-preservation.ts`'s `hasCapturedChildSpace` to decide
	 * whether this group's original `a:chOff`/`a:chExt` can be re-emitted
	 * verbatim (always true once captured, regardless of whether anything in
	 * the subtree has moved or resized - only its DIRECT children's
	 * `a:off`/`a:ext` are recomputed, via `invertChildIntoGroupSpace`, when
	 * something changed), instead of the normalized `chOff 0,0` / `chExt ==
	 * ext` space the writer falls back to when this is `undefined` (or
	 * degenerate). See `group-shape-geometry.ts`'s module doc for why a group
	 * needs two coordinate systems at all.
	 */
	chOffXEmu?: number | undefined;
	/** See {@link chOffXEmu}. */
	chOffYEmu?: number | undefined;
	/** See {@link chOffXEmu}. */
	chExtWidthEmu?: number | undefined;
	/** See {@link chOffXEmu}. */
	chExtHeightEmu?: number | undefined;
}

/**
 * A freehand ink / drawing stroke captured with a stylus or mouse.
 *
 * Ink strokes are stored as SVG path data strings. Each path may
 * have independent colour, width, and opacity.
 */
export interface InkPptxElement extends PptxElementBase {
	type: 'ink';
	/** SVG path data for ink strokes. */
	inkPaths: string[];
	/** Per-path stroke colours. */
	inkColors?: string[] | undefined;
	/** Per-path stroke widths. */
	inkWidths?: number[] | undefined;
	/** Per-path opacities (0-1). */
	inkOpacities?: number[] | undefined;
	/** Drawing tool used: pen, highlighter, or eraser. */
	inkTool?: 'pen' | 'highlighter' | 'eraser' | undefined;
	/**
	 * Per-path arrays of per-point pressure values (0-1).
	 *
	 * Each entry corresponds to the path at the same index in `inkPaths`.
	 * Each inner array contains one pressure value per sampled point along
	 * the stroke. When present, the renderer uses these values to produce
	 * variable-width strokes that reflect stylus/pen pressure.
	 */
	inkPointPressures?: number[][] | undefined;
	/** Per-path capture timestamps, in milliseconds on a common clock. */
	inkPointTimestamps?: number[][] | undefined;
	/**
	 * Per-path arrays of per-point pen-tilt lean direction (degrees, straight
	 * from `PointerEvent.tiltX` on supporting hardware).
	 *
	 * Each entry corresponds to the path at the same index in `inkPaths`, and
	 * is paired positionally with {@link inkPointTiltY}. Present only when at
	 * least one point in the stroke reported a genuinely non-zero tilt: a
	 * device that never reports tilt (a mouse, or a stylus with no tilt
	 * sensor) leaves both arrays absent, the same way `inkPointPressures` is
	 * omitted when pressure never varies. When present, the renderer converts
	 * the raw `(tiltX, tiltY)` vector into a lean angle + magnitude (see
	 * `pptx-viewer-shared`'s `tiltChannelsFromVectors`) and widens the stroke
	 * perpendicular to the lean direction, approximating a chisel-tip
	 * calligraphy nib.
	 */
	inkPointTiltX?: number[][] | undefined;
	/** Per-path, per-point pen-tilt lean direction (degrees), paired with {@link inkPointTiltX}. */
	inkPointTiltY?: number[][] | undefined;
	/** Unrecognised graphicFrame extLst extensions, captured verbatim for round-trip. */
	extensionXml?: PptxGraphicFrameExtension[] | undefined;
}

/**
 * A single ink stroke within a {@link ContentPartPptxElement}.
 */
export interface ContentPartInkStroke {
	path: string;
	color: string;
	width: number;
	opacity: number;
	/** Start and duration from InkML trace metadata, in milliseconds. */
	startTimestamp?: number;
	durationMs?: number;
	/**
	 * Per-point pressure values (0-1) for this stroke.
	 *
	 * When present, the renderer uses these values to produce
	 * variable-width strokes that reflect stylus/pen pressure.
	 */
	pressures?: number[] | undefined;
	/**
	 * Per-point pen-tilt lean direction (radians), decoded from the source
	 * InkML's `OTx`/`OTy` tilt-offset channels or its `AZIMUTH` channel.
	 *
	 * When present (paired with {@link tiltMagnitudes}), the renderer widens
	 * each point perpendicular to the lean direction, approximating a
	 * calligraphic (chisel-tip) nib. Absent when the source declared no tilt
	 * channel, in which case rendering is unaffected.
	 */
	tiltAngles?: number[] | undefined;
	/**
	 * Per-point pen-tilt strength (0 upright, 1 maximally leaned), paired with
	 * {@link tiltAngles}.
	 */
	tiltMagnitudes?: number[] | undefined;
	/**
	 * Which InkML channel pair {@link tiltAngles}/{@link tiltMagnitudes} were
	 * decoded from: `'azimuthAltitude'` when the source declared `AZIMUTH`
	 * (optionally paired with `ALTITUDE`); omitted (implying `OTx`/`OTy`, i.e.
	 * `'vector'`) otherwise, including for tilt this library itself captured
	 * from the Draw tool's `PointerEvent.tiltX`/`tiltY`.
	 *
	 * A save that has to rewrite this content part's InkML (see
	 * `inkml-content-part-writer.ts`) uses this to re-declare the SAME channel
	 * pair the file already used, rather than always converting to `OTx`/`OTy`;
	 * the rendered lean is identical either way; only the written channel
	 * NAMES differ.
	 */
	tiltEncoding?: 'vector' | 'azimuthAltitude' | undefined;
	/**
	 * Per-point timestamps (milliseconds), decoded from the source InkML
	 * trace's `T` channel when it declared one.
	 *
	 * Absent for the overwhelming majority of real decks: PowerPoint's own
	 * SaveAs output declares only `X`/`Y` and has no per-point time channel at
	 * all (see `contentpart-real-ink-roundtrip.test.ts`'s fixture), so this is
	 * populated only for InkML sources that genuinely author one (a captured
	 * digitizer session, OneNote, a Surface Hub export). When present, ink
	 * replay ("watch the ink get drawn") uses each stroke's own first/last
	 * timestamp to time its reveal instead of the fixed per-stroke cascade; see
	 * `pptx-viewer-shared`'s `render/ink-replay-timeline.ts`.
	 */
	pointTimestamps?: number[] | undefined;
}

/**
 * A content-part element wrapped in `mc:AlternateContent`.
 *
 * Typically contains ink strokes from modern PowerPoint pen/highlighter.
 */
export interface ContentPartPptxElement extends PptxElementBase {
	type: 'contentPart';
	/** Ink strokes contained in this content part. */
	inkStrokes?: ContentPartInkStroke[] | undefined;
	/** Package path of the related InkML part. */
	inkPartPath?: string | undefined;
	/** Parsed InkML root retained for unknown-node preservation on dirty save. */
	inkPartRawXml?: XmlObject | undefined;
}

/**
 * A Slide Zoom or Section Zoom element (PowerPoint Zoom Object).
 *
 * Zoom elements display a live thumbnail of the target slide and
 * navigate to it on click during presentation mode.
 *
 * @example
 * ```ts
 * const zoom: ZoomPptxElement = {
 *   type: "zoom",
 *   id: "zm_1", x: 300, y: 200, width: 200, height: 120,
 *   zoomType: "slide",
 *   targetSlideIndex: 5,
 * };
 * // => satisfies ZoomPptxElement
 * ```
 */
export interface ZoomPptxElement extends PptxElementBase, PptxImageProperties {
	type: 'zoom';
	/** Type of zoom: slide-level, section-level, or a multi-section summary. */
	zoomType: 'slide' | 'section' | 'summary';
	/** Zero-based index of the target slide. */
	targetSlideIndex: number;
	/** Whether the thumbnail is a slide preview or a custom cover. */
	zoomImageType?: 'preview' | 'cover' | undefined;
	/** Whether the thumbnail includes the destination slide background. */
	zoomShowBackground?: boolean | undefined;
	/** Section ID for section zoom. */
	targetSectionId?: string | undefined;
	/** Ordered section tiles in a Summary Zoom container. */
	summaryTargets?: SummaryZoomTarget[] | undefined;
	/** Layout mode authored on the Summary Zoom container. */
	summaryLayout?: 'grid' | 'fixed' | undefined;
	/**
	 * `zmPr/@returnToParent` (MS-PPTX `CT_ZoomObjectProperties`, shared by all
	 * three Zoom kinds): whether continuing forward from the destination slide
	 * during a live show returns to this Zoom's origin slide instead of
	 * advancing linearly through the deck. The schema declares `<xsd:attribute
	 * name="returnToParent" type="xsd:boolean" use="optional" default="true"/>`,
	 * so an ABSENT attribute means `true`, not `false`. This codebase still
	 * leaves the field `undefined` (rather than fabricating `true`) when the
	 * source XML omits it, so a round-trip never invents an attribute the
	 * source never had; a CONSUMER of this field (playback, e.g.
	 * `pptx-viewer-shared`'s `resolveZoomNavigationTarget`) must read
	 * `returnToParent !== false`, not `Boolean(returnToParent)`, to get the
	 * spec-correct effective value. For `zoomType: 'summary'`, this mirrors the
	 * first tile's own value; see {@link SummaryZoomTarget.returnToParent} for
	 * the per-tile value.
	 */
	returnToParent?: boolean | undefined;
	/**
	 * `zmPr/@transitionDur` (MS-PPTX `CT_ZoomObjectProperties`): the
	 * zoom-transition length in milliseconds. PowerPoint's own writer emits a
	 * unitless decimal-millisecond value for every OOXML `ST_UniversalTimeOffset`
	 * this codebase has observed (see the media-trim/-fade parsers), so this
	 * follows the same convention rather than the full TIMEOFFSET grammar's
	 * unit suffixes. Undefined uses the destination slide's own transition.
	 * For `zoomType: 'summary'`, this mirrors the first tile's own value; see
	 * {@link SummaryZoomTarget.transitionDurationMs} for the per-tile value.
	 */
	transitionDurationMs?: number | undefined;
}

/** A single section tile within a PowerPoint Summary Zoom container. */
export interface SummaryZoomTarget extends PptxImageProperties {
	sectionId: string;
	targetSlideIndex: number;
	zoomImageType?: ZoomPptxElement['zoomImageType'];
	zoomShowBackground?: boolean | undefined;
	x: number;
	y: number;
	width: number;
	height: number;
	title?: string | undefined;
	description?: string | undefined;
	offsetFactorX?: number | undefined;
	offsetFactorY?: number | undefined;
	scaleFactorX?: number | undefined;
	scaleFactorY?: number | undefined;
	/** This tile's own `zmPr/@returnToParent`; see {@link ZoomPptxElement.returnToParent}. */
	returnToParent?: boolean | undefined;
	/** This tile's own `zmPr/@transitionDur` in milliseconds; see {@link ZoomPptxElement.transitionDurationMs}. */
	transitionDurationMs?: number | undefined;
	rawXml?: XmlObject | undefined;
}

/**
 * A 3D model object embedded via `p16:model3D` inside an
 * `mc:AlternateContent` block (PowerPoint 365+).
 *
 * The element carries the path to the `.glb`/`.gltf` binary inside
 * the ZIP and a poster/fallback image for rendering in viewers that
 * do not support interactive 3D.
 */
export interface Model3DPptxElement extends PptxElementBase, PptxImageProperties {
	type: 'model3d';
	/** Path to the 3D model file inside the ZIP. */
	modelPath?: string | undefined;
	/** Base64 data URL of the 3D model binary. */
	modelData?: string | undefined;
	/** MIME type of the model (e.g. "model/gltf-binary"). */
	modelMimeType?: string | undefined;
	/** Poster/preview image shown when 3D rendering is unavailable. */
	posterImage?: string | undefined;
	/**
	 * Authored camera, model transform and lights (`am3d:camera`, `am3d:trans`,
	 * lights). Read-only view: the XML round-trips via `rawXml`.
	 */
	scene?: Model3DSceneData | undefined;
	/** Unrecognised graphicFrame extLst extensions, captured verbatim for round-trip. */
	extensionXml?: PptxGraphicFrameExtension[] | undefined;
}

/** An element whose type is not recognised by the parser. */
export interface UnknownPptxElement extends PptxElementBase {
	type: 'unknown';
	/** Unrecognised graphicFrame extLst extensions, captured verbatim for round-trip. */
	extensionXml?: PptxGraphicFrameExtension[] | undefined;
}

// ==========================================================================
// Discriminated union
// ==========================================================================

/**
 * A single element on a PPTX slide.
 *
 * This is a **discriminated union**: narrow on `element.type` to access
 * variant-specific properties like `imageData` (image/picture), `pathData`
 * (shape), or `textSegments` (text/shape).
 */
export type PptxElement =
	| TextPptxElement
	| ShapePptxElement
	| ConnectorPptxElement
	| ImagePptxElement
	| PicturePptxElement
	| TablePptxElement
	| ChartPptxElement
	| SmartArtPptxElement
	| OlePptxElement
	| MediaPptxElement
	| GroupPptxElement
	| InkPptxElement
	| ContentPartPptxElement
	| ZoomPptxElement
	| Model3DPptxElement
	| UnknownPptxElement;

/**
 * Discriminant values for the `type` field on {@link PptxElement}.
 *
 * DERIVED from the union on purpose. This alias used to be a hand-written list
 * of string literals living in `types/common.ts`, and it drifted: two element
 * types (`contentPart` and `model3d`) were added to {@link PptxElement} without
 * being added to the list, so every consumer that keyed a registry or a switch
 * off the alias was silently blind to ink content parts and 3D models while
 * still type-checking. Deriving it makes that class of drift impossible, at the
 * cost of nothing: the resolved type is identical.
 *
 * Narrow on this type to access variant-specific properties.
 *
 * @example
 * ```ts
 * function isImage(el: PptxElement): el is ImagePptxElement {
 *   return el.type === "image";
 * }
 * // => type guard narrowing PptxElement to ImagePptxElement
 * ```
 */
export type PptxElementType = PptxElement['type'];

// ==========================================================================
// Utility type aliases (for function signatures that accept subsets)
// ==========================================================================

/** Elements that can contain text content (text boxes, shapes, and connectors). */
export type PptxElementWithText = TextPptxElement | ShapePptxElement | ConnectorPptxElement;

/** Elements that carry shape styling (fill, stroke, geometry). */
export type PptxElementWithShapeStyle =
	| TextPptxElement
	| ShapePptxElement
	| ConnectorPptxElement
	| ImagePptxElement
	| PicturePptxElement;

/** Elements that hold raster image data. */
export type PptxImageLikeElement = ImagePptxElement | PicturePptxElement;
