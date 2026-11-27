import type { InkStroke, InkXmlObject } from './types';
import { resolveInkClock, strokeStart, timestampOffset } from './inkml-write-clock';

/**
 * The writer half of `inkml-content-part.ts` (split out to keep both files
 * under this repo's file-size guideline): build schema-shaped InkML XML from
 * an in-memory stroke list, in this project's own authored dialect (verified
 * to open in real PowerPoint; see the structural notes on {@link buildInkMlContent}).
 *
 * @module drawingml/inkml/inkml-content-part-writer
 */

const INKML_NAMESPACE = 'http://www.w3.org/2003/InkML';
const METADATA_NAMESPACE = 'https://pptx-viewer.dev/inkml/metadata';

/** Decoded per-point pen-tilt lean, mirrored from `inkml-trace-decode.ts`'s `TiltChannels`. */
interface TiltChannels {
	angles: readonly number[];
	magnitudes: readonly number[];
}

/**
 * Which tilt channel pair (if any) the whole part's shared `traceFormat`
 * declares. `'azimuthAltitude'` is used only when EVERY tilt-carrying stroke
 * was itself decoded from `AZIMUTH`/`ALTITUDE` (see
 * {@link InkStroke.tiltEncoding}); a mixed part (or one containing
 * any stroke whose tilt has no recorded encoding, i.e. this library's own
 * `OTx`/`OTy`-equivalent capture) falls back to `'vector'`, since a single
 * `traceFormat` cannot declare two competing tilt encodings for different
 * traces. Either way the rendered lean is identical: only the written channel
 * NAMES differ.
 */
type TiltMode = 'none' | 'vector' | 'azimuthAltitude';

/** Decide the whole part's tilt channel declaration from its strokes. */
function resolveTiltMode(strokes: readonly InkStroke[]): TiltMode {
	const tiltStrokes = strokes.filter(hasStrokeTilt);
	if (tiltStrokes.length === 0) {
		return 'none';
	}
	return tiltStrokes.every((stroke) => stroke.tiltEncoding === 'azimuthAltitude')
		? 'azimuthAltitude'
		: 'vector';
}

/** Build schema-shaped InkML while retaining unknown nodes from a loaded part. */
export function buildInkMlContent(
	strokes: readonly InkStroke[],
	rawXml?: InkXmlObject,
	space?: { width: number; height: number },
): InkXmlObject {
	const data = rawXml ? { ...rawXml } : {};
	// A loaded PowerPoint part is keyed `inkml:ink`, not `ink:ink`. Missing that
	// wrote a SECOND root element beside PowerPoint's, producing an XML part with
	// two roots that no consumer can read.
	const existingKey = Object.keys(data).find((key) => localNameOf(key) === 'ink');
	const existingRoot = existingKey ? (data[existingKey] as InkXmlObject | undefined) : undefined;
	const root: InkXmlObject = existingRoot ? { ...existingRoot } : {};
	deleteStaleInkChildren(root);
	root['@_xmlns:ink'] = INKML_NAMESPACE;
	root['@_xmlns:pva'] = METADATA_NAMESPACE;
	// The box the authored paths are drawn in, so a reader stretches them onto a resized drawing.
	if (space) {
		root['@_pva:spaceWidth'] = String(space.width);
		root['@_pva:spaceHeight'] = String(space.height);
	} else {
		delete root['@_pva:spaceWidth'];
		delete root['@_pva:spaceHeight'];
	}
	// Only declare (and author per-point values for) a tilt channel pair when
	// at least one stroke actually carries tilt data. A document with no tilt
	// anywhere must serialise byte-identically to before this feature existed:
	// no extra channels, no extra trailing columns.
	const tiltMode = resolveTiltMode(strokes);
	const clock = resolveInkClock(strokes);
	const mixed = clock !== undefined && strokes.some((stroke) => strokeStart(stroke) === undefined);
	// Verified against real PowerPoint COM behaviour: a plain `<ink:traceFormat>`
	// / `<ink:brush>` / `id="..."` part (this project's ORIGINAL authored
	// dialect) passes this project's own lenient reader and internal schema
	// validator, but real PowerPoint's own InkML parser rejects it as
	// "corrupted and unreadable" (0x80070570). Real PowerPoint requires:
	//   - `traceFormat` nested inside `definitions/context/inkSource`, not a
	//     direct child of the root.
	//   - `brush` nested inside `definitions` (a sibling of `context`), not a
	//     direct child of the root.
	//   - Every identifiable element (`context`, `inkSource`, `brush`) keyed by
	//     the InkML/XML spec's `xml:id`, NOT a bare `id` attribute: `id="..."`
	//     alone is schema-legal enough for this project's own reader but real
	//     PowerPoint's parser rejects the whole part.
	//   - Each `<trace>` carries BOTH `contextRef` and `brushRef`; `brushRef`
	//     alone (this project's original authored form) is rejected too.
	// The compact difference-encoding PowerPoint itself writes (`100
	// 200,'40'46,"0"-5`) is NOT required: plain per-point decimal channel
	// values (already this project's own dialect) open fine once the
	// structural requirements above are met.
	root['ink:definitions'] = {
		'ink:context': mixed
			? [buildContext('ctx0', tiltMode, clock), buildContext('ctx1', tiltMode)]
			: buildContext('ctx0', tiltMode, clock),
		'ink:brush': strokes.map((stroke, index) => ({
			'@_xml:id': `brush${index + 1}`,
			'ink:brushProperty': brushProperties(stroke),
		})),
	};
	root['ink:trace'] = strokes.map((stroke, index) => {
		const start = strokeStart(stroke);
		return {
			'@_contextRef': mixed && start === undefined ? '#ctx1' : '#ctx0',
			'@_brushRef': `#brush${index + 1}`,
			'@_pva:path': stroke.path,
			...(clock && start !== undefined
				? { '@_timestampRef': '#clock', '@_timeOffset': timestampOffset(start, clock.base) }
				: {}),
			...(Number.isFinite(stroke.durationMs) ? { '@_duration': String(stroke.durationMs) } : {}),
			...pathToTrace(stroke, tiltMode, clock?.base),
		};
	});
	// The rewritten root replaces whatever prefix the source used, and the
	// source's own definitions/traces must not survive beside it. Also drop a
	// root-level `ink:traceFormat`/`ink:brush` a PRE-fix version of this writer
	// left directly on the root (real PowerPoint's parser requires both nested
	// under `ink:definitions`, moved above): those are stale duplicates once
	// this rebuild runs, not additional content to keep.
	if (existingKey) {
		delete data[existingKey];
	}
	delete data['ink'];
	data['ink:ink'] = root;
	return data;
}

/**
 * A brush as Word, Excel and PowerPoint 16 read it, checked by opening saved files in each: the
 * size as `width` and `height` in centimetres (Word draws a unitless width as a grey
 * hairline, ignoring the colour too), the colour, and InkML `transparency` (0 opaque to 255
 * invisible), which the applications honour; `opacity` keeps the exact value for this library's
 * own reader.
 *
 * Earlier releases wrote the width without units (pixels, as this reader assumed), so PowerPoint
 * drew re-saved ink as thin grey lines. Parts written that way still read back at their old
 * pixel widths (a unitless value is pixels), and saving them again writes centimetres, so ink a
 * deck re-saves now keeps its colour and thickness in PowerPoint, Word and Excel. A reader of
 * these parts that assumed unitless pixel widths must now honour `units="cm"`.
 */
function brushProperties(stroke: InkStroke): InkXmlObject[] {
	const size = String((stroke.width * 2.54) / 96);
	const opacity = Math.max(0, Math.min(1, stroke.opacity));
	const extras = new Map((stroke.brushExtras ?? []).map((extra) => [extra.name, extra]));
	// Office's highlighter: a rectangular tip that masks (multiplies) what is under it.
	if (stroke.tool === 'highlighter') {
		if (!extras.has('tip')) extras.set('tip', { name: 'tip', value: 'rectangle' });
		if (!extras.has('rasterOp')) extras.set('rasterOp', { name: 'rasterOp', value: 'maskPen' });
	}
	const height = extras.get('height');
	extras.delete('height');
	return [
		{ '@_name': 'width', '@_value': size, '@_units': 'cm' },
		height
			? {
					'@_name': 'height',
					'@_value': height.value,
					...(height.units ? { '@_units': height.units } : {}),
				}
			: { '@_name': 'height', '@_value': size, '@_units': 'cm' },
		{ '@_name': 'color', '@_value': stroke.color },
		...(opacity < 1
			? [{ '@_name': 'transparency', '@_value': String(Math.round((1 - opacity) * 255)) }]
			: []),
		{ '@_name': 'opacity', '@_value': String(stroke.opacity) },
		// The source part's own properties this model does not hold, as they were written.
		...[...extras.values()].map((extra) => ({
			'@_name': extra.name,
			'@_value': extra.value,
			...(extra.units ? { '@_units': extra.units } : {}),
		})),
	];
}

/** Drop the source part's own trace/brush/definition nodes after a rewrite. */
function deleteStaleInkChildren(root: InkXmlObject): void {
	for (const key of Object.keys(root)) {
		if (key.startsWith('@_') || key === '#text') {
			continue;
		}
		const local = localNameOf(key);
		if (
			local === 'trace' ||
			local === 'brush' ||
			local === 'traceGroup' ||
			local === 'traceFormat' ||
			local === 'definitions'
		) {
			delete root[key];
		}
	}
}

function localNameOf(key: string): string {
	const colon = key.indexOf(':');
	return colon >= 0 ? key.slice(colon + 1) : key;
}

function buildContext(
	id: string,
	mode: TiltMode,
	clock?: ReturnType<typeof resolveInkClock>,
): InkXmlObject {
	return {
		'@_xml:id': id,
		'ink:inkSource': {
			'@_xml:id': `inkSrc${id.slice(3)}`,
			'ink:traceFormat': { 'ink:channel': buildTraceFormatChannels(mode) },
		},
		...(clock
			? { 'ink:timestamp': { '@_xml:id': 'clock', '@_timeString': clock.timeString } }
			: {}),
	};
}

/** The three (no tilt), five (vector), or five (azimuth/altitude) `<ink:channel>` declarations for the shared traceFormat. */
function buildTraceFormatChannels(mode: TiltMode): InkXmlObject[] {
	const channels: InkXmlObject[] = [
		{ '@_name': 'X', '@_type': 'decimal' },
		{ '@_name': 'Y', '@_type': 'decimal' },
		{ '@_name': 'F', '@_type': 'decimal', '@_min': '0', '@_max': '1' },
	];
	if (mode === 'vector') {
		// `OTx`/`OTy`: the InkML-conventional tilt-offset channel pair the
		// reader's `pointsToTilt` already decodes (see `inkml-trace-decode.ts`).
		// Declared as unbounded decimals (no min/max): the values are an
		// arbitrary-scale direction vector, not a normalised 0-1 reading like
		// `F`, matching how a real digitizer's own tilt-offset channel is
		// typically declared.
		channels.push(
			{ '@_name': 'OTx', '@_type': 'decimal' },
			{ '@_name': 'OTy', '@_type': 'decimal' },
		);
	} else if (mode === 'azimuthAltitude') {
		// Re-declares the SAME pair a loaded stroke's tilt was originally
		// decoded from (see `tiltEncoding`), so a save that has to rewrite this
		// part (an edit elsewhere in it) does not silently convert every
		// AZIMUTH/ALTITUDE-authored stroke to OTx/OTy. Degrees, unbounded (a
		// negative or >360 azimuth reads back identically; see `tiltOffsetAt`'s
		// counterpart `azimuthAltitudeAt`).
		channels.push(
			{ '@_name': 'AZIMUTH', '@_type': 'decimal' },
			{ '@_name': 'ALTITUDE', '@_type': 'decimal' },
		);
	}
	return channels;
}

/** True when a stroke carries tilt data with any point genuinely leaned (non-zero magnitude). */
function hasStrokeTilt(stroke: InkStroke): boolean {
	return Boolean(
		stroke.tiltAngles?.length && stroke.tiltMagnitudes?.some((magnitude) => magnitude > 1e-4),
	);
}

/** Bundle a stroke's tilt arrays for {@link pathToTrace}, or `undefined` when it has none. */
function tiltOf(stroke: InkStroke): TiltChannels | undefined {
	return stroke.tiltAngles && stroke.tiltMagnitudes
		? { angles: stroke.tiltAngles, magnitudes: stroke.tiltMagnitudes }
		: undefined;
}

function pathToTrace(stroke: InkStroke, mode: TiltMode, base: number | undefined): InkXmlObject {
	const { path, pressures, pointTimestamps } = stroke;
	const tilt = tiltOf(stroke);
	const points = [...path.matchAll(/[ML]\s*(?<x>[\d.eE+-]+)[,\s]+(?<y>[\d.eE+-]+)/giu)];
	if (points.length === 0) return { '#text': path };

	// PowerPoint 16 rejects T in the trace format. Keep the same integer
	// offsets in our metadata and leave the native coordinate channels readable.
	const offsets: string[] = [];
	const timed = base !== undefined && pointTimestamps?.some(Number.isFinite);
	let sample = strokeStart(stroke) ?? base ?? 0;
	const text = points
		.map((point, index) => {
			if (timed) {
				const recorded = pointTimestamps?.[index];
				if (typeof recorded === 'number' && Number.isFinite(recorded)) sample = recorded;
				offsets.push(timestampOffset(sample, base));
			}
			const pressure = Math.max(0, Math.min(1, pressures?.[index] ?? 0.5));
			const xyf = `${point.groups?.x} ${point.groups?.y} ${pressure}`;
			if (mode === 'none') return xyf;
			const { a, b } =
				mode === 'azimuthAltitude' ? azimuthAltitudeAt(tilt, index) : tiltOffsetAt(tilt, index);
			return `${xyf} ${a} ${b}`;
		})
		.join(', ');
	return { '#text': text, ...(offsets.length ? { '@_pva:sampleTimes': offsets.join(',') } : {}) };
}

/**
 * Inverse of the reader's `tiltFromVector`/`tiltChannelsFromXY`: a point's
 * `(angle, magnitude)` reads back to itself exactly when `magnitude` is
 * already peak-normalised across the stroke (max 1), which is exactly what
 * `pointsToTilt` (and this project's own `tiltChannelsFromVectors` capture
 * helper) always produce, so `ox = magnitude * cos(angle)`,
 * `oy = magnitude * sin(angle)` is a lossless round-trip for any stroke this
 * project itself authored. A point with no tilt data, or a magnitude of 0,
 * authors `"0 0"` (upright / no lean), matching how an absent tilt channel
 * already renders (a plain circle, not an ellipse).
 */
function tiltOffsetAt(tilt: TiltChannels | undefined, index: number): { a: string; b: string } {
	const angle = tilt?.angles[index];
	const magnitude = tilt?.magnitudes[index];
	if (angle === undefined || magnitude === undefined || magnitude <= 0) {
		return { a: '0', b: '0' };
	}
	return {
		a: roundTiltComponent(magnitude * Math.cos(angle)),
		b: roundTiltComponent(magnitude * Math.sin(angle)),
	};
}

/**
 * Inverse of the reader's `tiltFromAzimuthAltitude`: `azimuth = angle` (already
 * radians, converted straight back to degrees; a negative value round-trips
 * fine since the reader re-derives radians from the same number rather than
 * expecting a normalised 0-360 range) and `altitude = (1 - magnitude) * 90`
 * (upright/magnitude-0 reads back as altitude 90, matching the reader's own
 * "no altitude channel" default). A point with no tilt data authors `"0 90"`
 * (azimuth 0, fully upright), the AZIMUTH/ALTITUDE equivalent of
 * {@link tiltOffsetAt}'s `"0 0"`.
 */
function azimuthAltitudeAt(
	tilt: TiltChannels | undefined,
	index: number,
): { a: string; b: string } {
	const angle = tilt?.angles[index];
	const magnitude = tilt?.magnitudes[index];
	if (angle === undefined || magnitude === undefined || magnitude <= 0) {
		return { a: '0', b: '90' };
	}
	return {
		a: roundTiltComponent((angle * 180) / Math.PI),
		b: roundTiltComponent(90 * (1 - Math.min(1, magnitude))),
	};
}

/** Round a written tilt channel component to 6 decimal places (compact, imperceptibly lossy). */
function roundTiltComponent(value: number): string {
	return String(Math.round(value * 1e6) / 1e6);
}
