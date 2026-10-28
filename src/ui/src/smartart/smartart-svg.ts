import type { DiagramPictureResolver } from './smartart-picture';
import { hexToRgbChannels } from 'ooxml-core/color';
import { resolveDrawingColor } from 'ooxml-core/diagram';
import type { DiagramDrawing } from 'ooxml-core/diagram';
import type { DrawingColor } from 'ooxml-core/drawingml';

import { diagramShapeElement } from './smartart-geometry';
import { diagramFill } from './smartart-paint';
import { diagramTextElement } from './smartart-text';
import { diagramEffects } from './smartart-effects';

const SVG_NS = 'http://www.w3.org/2000/svg';
const EMU_PER_PX = 9525;

export interface SmartArtRenderReport {
	shapeCount: number;
	/** Preset geometries drawn as a plain rectangle because no outline is available here. */
	approximatedGeometries: string[];
	/** Shapes that carry 3D scene data: drawn flat, the 3D is kept only in the source part. */
	flattened3d: number;
	/** Unsupported fill kinds and approximated path-gradient geometries. */
	approximatedFills: string[];
	/** Text layouts that require unavailable fonts or unsupported typography. */
	approximatedText?: string[];
	unappliedEffects?: string[];
	/** DrawingML transforms the shared resolver cannot apply yet. */
	unappliedColorTransforms?: string[];
}

/** Theme colours the host resolved (`accent1` -> `#4472c4`); unresolved scheme colours are grey. */
export type SchemeColors = Readonly<Record<string, string>>;
export interface SchemeFonts {
	major?: string;
	minor?: string;
}

export function colorToCss(
	color: DrawingColor | undefined,
	scheme: SchemeColors,
	report?: SmartArtRenderReport,
): string | undefined {
	if (!color) return undefined;
	const resolved = resolveDrawingColor(color, { scheme: (name) => scheme[name] });
	const unapplied = resolved?.unapplied ?? color.transforms.map((transform) => transform.name);
	if (report && unapplied.length)
		report.unappliedColorTransforms = [
			...new Set([...(report.unappliedColorTransforms ?? []), ...unapplied]),
		];
	if (resolved) {
		if (resolved.alpha === 1) return resolved.hex;
		const rgb = hexToRgbChannels(resolved.hex)!;
		return `rgba(${rgb.r}, ${rgb.g}, ${rgb.b}, ${resolved.alpha})`;
	}
	if (color.kind === 'preset')
		return /^[a-z]+$/i.test(color.value) ? color.value.toLowerCase() : '#808080';
	if (color.kind === 'system' && !color.fallback) return undefined;
	return '#9ca3af';
}

/** Render a core `DiagramDrawing` into an `<svg>`: pure DOM construction, no markup injection. */
export function renderDiagramDrawing(
	doc: Document,
	drawing: DiagramDrawing,
	scheme: SchemeColors = {},
	fonts: SchemeFonts = {},
	resolvePicture?: DiagramPictureResolver,
): { svg: SVGSVGElement; report: SmartArtRenderReport } {
	const report: SmartArtRenderReport = {
		shapeCount: drawing.shapes.length,
		approximatedGeometries: [],
		flattened3d: 0,
		approximatedFills: [],
	};
	const svg = doc.createElementNS(SVG_NS, 'svg');
	const defs = doc.createElementNS(SVG_NS, 'defs');
	svg.append(defs);
	let right = 0;
	let bottom = 0;
	for (const shape of drawing.shapes) {
		const { frame } = shape;
		right = Math.max(right, (frame.x + frame.width) / EMU_PER_PX);
		bottom = Math.max(bottom, (frame.y + frame.height) / EMU_PER_PX);
		const group = doc.createElementNS(SVG_NS, 'g');
		group.dataset.modelId = shape.modelId;
		const cx = frame.width / EMU_PER_PX / 2;
		const cy = frame.height / EMU_PER_PX / 2;
		const fx = shape.flipHorizontal ? -1 : 1;
		const fy = shape.flipVertical ? -1 : 1;
		const transform = [`translate(${frame.x / EMU_PER_PX} ${frame.y / EMU_PER_PX})`];
		if (shape.rotation) transform.push(`rotate(${shape.rotation} ${cx} ${cy})`);
		if (fx !== 1 || fy !== 1) {
			transform.push(`translate(${cx} ${cy}) scale(${fx} ${fy}) translate(${-cx} ${-cy})`);
		}
		group.setAttribute('transform', transform.join(' '));
		const body = diagramShapeElement(doc, shape, report.approximatedGeometries);
		body.setAttribute(
			'fill',
			diagramFill(doc, defs, shape.fill, scheme, report, drawing.pictureParts, resolvePicture, {
				width: frame.width / EMU_PER_PX,
				height: frame.height / EMU_PER_PX,
				...(shape.rotation !== undefined ? { rotation: shape.rotation } : {}),
				...(shape.flipHorizontal ? { flipHorizontal: true } : {}),
				...(shape.flipVertical ? { flipVertical: true } : {}),
			}),
		);
		diagramEffects(doc, defs, body, shape, scheme, report);
		const stroke =
			shape.line?.fill && shape.line.widthEmu !== 0
				? diagramFill(
						doc,
						defs,
						shape.line.fill,
						scheme,
						report,
						drawing.pictureParts,
						resolvePicture,
					)
				: undefined;
		body.setAttribute('stroke', stroke ?? 'none');
		if (stroke)
			body.setAttribute(
				'stroke-width',
				String(Math.max(0.5, (shape.line?.widthEmu ?? 9525) / EMU_PER_PX)),
			);
		group.append(body);
		if (shape.has3d) {
			report.flattened3d += 1;
			group.setAttribute('data-flattened-3d', '');
		}
		if (shape.text?.text) group.append(diagramTextElement(doc, shape, scheme, fonts, report));
		svg.append(group);
	}
	svg.setAttribute('viewBox', `0 0 ${Math.max(right, 1)} ${Math.max(bottom, 1)}`);
	// The viewBox is the shapes' frames, and a centred outline straddles them: a
	// node on the drawing's edge lost the outer half of its line to the
	// viewport clip an inline <svg> applies by default.
	svg.setAttribute('overflow', 'visible');
	return { svg, report };
}
