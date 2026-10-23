import JSZip from 'jszip';
import { readFileSync, readdirSync, statSync } from 'node:fs';
import path from 'node:path';
import { beforeAll, describe, expect, it } from 'vitest';

import {
	parseDiagramDrawing,
	readDiagramDrawingShapes,
	resolveDrawingColor,
	type DiagramColor,
	type DiagramDrawingShape,
	type DiagramShapeView,
} from '../../../diagram/index';
import { PptxRuntimeDependencyFactory } from '../../core/core/factories/PptxRuntimeDependencyFactory';
import { PptxHandlerRuntime } from '../../core/core/runtime/PptxHandlerRuntimeImplementation';
import type { PptxSmartArtDrawingShape, XmlObject } from '../../core/types';

/**
 * SmartArt drawing-shape parity (docs/agnostic-core-plan.md, SmartArt step S2 / plan step 4):
 * every cached drawing part (`dsp:drawing`) of the committed pptx, docx and xlsx fixtures is read
 * by BOTH the pptx runtime's `parseDrawingShape` (object tree, `this.parseColor`) and the neutral
 * DOM reader (`parseDiagramDrawing` + `resolveDiagramDrawingShape`), and every field of the
 * neutral `DiagramShapeView` must agree.
 *
 * The neutral side resolves colours through the shared DrawingML pipeline over a theme whose slots
 * are the runtime's own (its default scheme map), so the transforms (`lumMod`, `shade`, `alpha`...)
 * are compared, not the theme. Group shapes (`dsp:grpSp`) are compared at top level only, since
 * the pptx reader does not descend into groups.
 *
 * One known divergence is counted rather than compared: a colour that repeats a transform (two
 * `a:tint` children, as some 3D quick styles cache their gradient stops). The object tree groups
 * same-name siblings into an array the pptx colour codec does not read, so pptx skips those
 * transforms; the neutral resolver applies them. Which one PowerPoint matches is not established.
 */

const CORE = path.resolve(__dirname, '../../..');

class ParityRuntime extends PptxHandlerRuntime {
	public shape(sp: XmlObject, index: number): PptxSmartArtDrawingShape | null {
		return this.parseDrawingShape(sp, index, 9525);
	}
	public scheme(name: string): string | undefined {
		return this.parseColor({ 'a:schemeClr': { '@_val': name } }) ?? undefined;
	}
}

const runtime = new ParityRuntime();
const parser = new PptxRuntimeDependencyFactory().createParser();
const theme = { scheme: (name: string) => runtime.scheme(name) };
const resolveColor = (color: Parameters<typeof resolveDrawingColor>[0]) =>
	resolveDrawingColor(color, theme);

function packages(): string[] {
	const found: string[] = [];
	const walk = (dir: string) => {
		for (const name of readdirSync(dir).sort()) {
			const full = path.join(dir, name);
			if (statSync(full).isDirectory()) walk(full);
			else if (name.endsWith('.pptx')) found.push(full);
		}
	};
	walk(path.join(CORE, 'pptx/__tests__/fixtures'));
	return [
		...found,
		path.join(CORE, 'docx/__fixtures__/smartart.docx'),
		path.join(CORE, 'xlsx/__fixtures__/excel-smartart.xlsx'),
	];
}

const DRAWING_PART = /(?:^|\/)diagrams\/drawing[^/]*\.xml$/u;

function local(key: string): string {
	return key.split(':').pop() ?? key;
}
function entry(node: XmlObject | undefined, name: string): unknown {
	if (!node) return undefined;
	const key = Object.keys(node).find((candidate) => local(candidate) === name);
	return key ? node[key] : undefined;
}
function list(value: unknown): XmlObject[] {
	if (value === '') return [{}];
	return (Array.isArray(value) ? value : value ? [value] : []) as XmlObject[];
}

/** The fields of the pptx shape that the neutral view also models, normalised. */
const VIEW_FIELDS = [
	'id',
	'shapeType',
	'x',
	'y',
	'width',
	'height',
	'rotation',
	'flipHorizontal',
	'flipVertical',
	'shapeAdjustments',
	'fillNone',
	'fillColor',
	'fillOpacity',
	'fillGradientStops',
	'fillGradientType',
	'fillGradientAngle',
	'fillPatternPreset',
	'fillPatternForegroundColor',
	'fillPatternBackgroundColor',
	'fillBlipEmbedId',
	'hasShadow',
	'shadowColor',
	'strokeColor',
	'strokeWidth',
	'text',
	'fontSize',
	'fontColor',
	'fontFamily',
	'fontWeight',
	'fontStyle',
	'lineHeight',
	'lineHeightRatio',
	'lineSpacingAfter',
	'lineSpacingAfterRatio',
	'textInsetLeft',
	'textInsetTop',
	'textInsetRight',
	'textInsetBottom',
	'textVerticalAnchor',
	'textFrameX',
	'textFrameY',
	'textFrameWidth',
	'textFrameHeight',
] as const satisfies readonly (keyof DiagramShapeView)[];

function normalise(shape: DiagramShapeView): Record<string, unknown> {
	const out: Record<string, unknown> = {};
	for (const field of VIEW_FIELDS) {
		let value: unknown = shape[field];
		if (value === false || value === undefined || value === null) continue;
		if (field === 'rotation' && value === 0) continue;
		if (typeof value === 'number') value = Math.round(value * 1000) / 1000;
		if (field === 'fillGradientStops')
			value = (value as DiagramShapeView['fillGradientStops'])!.map((stop) => ({
				...stop,
				position: Math.round(stop.position * 1000) / 1000,
			}));
		if (field === 'shapeAdjustments' && Object.keys(value as object).length === 0) continue;
		out[field] = value;
	}
	return out;
}

interface Difference {
	label: string;
	field: string;
	pptx: unknown;
	neutral: unknown;
}

const differences: Difference[] = [];
let shapesCompared = 0;
let partsCompared = 0;
let fieldsCompared = 0;
let repeatedTransformFields = 0;

const REPEATED_TRANSFORM_FIELDS = new Set(['fillColor', 'fillGradientStops']);

const colorRepeats = (color: DiagramColor | undefined) =>
	!!color &&
	new Set(color.transforms.map((transform) => transform.name)).size !== color.transforms.length;

/** The shape's fill holds a colour with a repeated transform (see the module comment). */
function repeatsTransform(shape: DiagramDrawingShape): boolean {
	const fill = shape.fill;
	if (fill?.kind === 'solid') return colorRepeats(fill.color);
	if (fill?.kind === 'gradient') return fill.stops.some((stop) => colorRepeats(stop.color));
	return false;
}

beforeAll(async () => {
	for (const file of packages()) {
		let zip: JSZip;
		try {
			zip = await JSZip.loadAsync(readFileSync(file));
		} catch {
			continue;
		}
		for (const part of Object.keys(zip.files)
			.filter((name) => DRAWING_PART.test(name))
			.sort()) {
			const xml = await zip.file(part)!.async('string');
			const tree = parser.parse(xml) as XmlObject;
			const drawing = entry(tree, 'drawing') as XmlObject | undefined;
			const spTree = entry(drawing ?? tree, 'spTree') as XmlObject | undefined;
			const pptxShapes = list(entry(spTree, 'sp'))
				.map((sp, index) => runtime.shape(sp, index))
				.filter((shape): shape is PptxSmartArtDrawingShape => shape !== null);
			const repeated = new Set(
				parseDiagramDrawing(xml, part)
					.shapes.filter(repeatsTransform)
					.map((shape) => shape.modelId),
			);
			const topLevel = new Set(
				list(entry(spTree, 'sp')).map((sp, index) =>
					String(
						sp['@_modelId'] ||
							(entry(entry(sp, 'nvSpPr') as XmlObject, 'cNvPr') as XmlObject | undefined)?.[
								'@_id'
							] ||
							`dsp-${index}`,
					),
				),
			);
			const neutralShapes = readDiagramDrawingShapes(xml, { resolveColor }, part).shapes.filter(
				(shape) => topLevel.has(shape.id),
			);
			partsCompared++;
			const label = `${path.relative(CORE, file)}!${part}`;
			if (pptxShapes.length !== neutralShapes.length) {
				differences.push({
					label,
					field: 'count',
					pptx: pptxShapes.length,
					neutral: neutralShapes.length,
				});
				continue;
			}
			pptxShapes.forEach((pptxShape, index) => {
				shapesCompared++;
				const left = normalise(pptxShape);
				const right = normalise(neutralShapes[index]!);
				for (const field of new Set([...Object.keys(left), ...Object.keys(right)])) {
					if (repeated.has(pptxShape.id) && REPEATED_TRANSFORM_FIELDS.has(field)) {
						repeatedTransformFields++;
						continue;
					}
					fieldsCompared++;
					if (JSON.stringify(left[field]) !== JSON.stringify(right[field]))
						differences.push({
							label: `${label}#${index}`,
							field,
							pptx: left[field],
							neutral: right[field],
						});
				}
			});
		}
	}
}, 240_000);

describe('SmartArt drawing shapes: pptx runtime vs neutral DOM reader', () => {
	it('compares a real corpus', () => {
		expect(partsCompared).toBeGreaterThan(300);
		expect(shapesCompared).toBeGreaterThan(2000);
		expect(fieldsCompared).toBeGreaterThan(20000);
		// The repeated-transform divergence stays confined to a few 3D gradient shapes.
		expect(repeatedTransformFields).toBeLessThan(200);
	});

	it('agrees on every shape of every committed drawing part', () => {
		expect(differences.slice(0, 20)).toEqual([]);
	});
});
