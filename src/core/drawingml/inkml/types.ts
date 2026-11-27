/**
 * The format-neutral InkML model shared by PowerPoint (`p:contentPart`), Word (`w14:contentPart`)
 * and Excel (`xdr:contentPart`): one stroke list read from and written to an `application/inkml+xml`
 * part, and the object tree the reader and writer work on.
 *
 * @module drawingml/inkml/types
 */

/**
 * An XML element as an object tree: attributes under `@_`-prefixed keys, text under `#text`, and
 * children keyed by their qualified name (one object, an array for repeated elements, or a bare
 * string for an element that holds only text). PowerPoint's parser produces this shape directly;
 * Word and Excel convert their parsed DOM with `inkObjectFromXml`.
 */
export interface InkXmlObject {
	[attr: `@_${string}`]: string | undefined;
	'#text'?: string;
	[child: string]: InkXmlObject | InkXmlObject[] | string | undefined;
}

/**
 * One ink stroke, in the coordinate space of the drawing that holds it (CSS pixels of its box).
 */
export interface InkStroke {
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
	/**
	 * The pen that drew it, when known: this library's own pens, or `highlighter` for a brush with
	 * InkML `rasterOp="maskPen"` (Office's highlighter). The writer gives a highlighter Office's
	 * highlighter brush (`tip="rectangle"`, `rasterOp="maskPen"`).
	 */
	tool?: 'pen' | 'pencil' | 'highlighter' | 'actionPen' | undefined;
	/**
	 * Brush properties of the source part this model does not hold (`tip`, `rasterOp`,
	 * `ignorePressure`, `fitToCurve`, a `height` other than the width...), written back as read.
	 */
	brushExtras?: InkBrushProperty[] | undefined;
}

/** One InkML `brushProperty`, as written. */
export interface InkBrushProperty {
	name: string;
	value: string;
	units?: string | undefined;
}
