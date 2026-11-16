import { beforeAll, describe, expect, it } from 'vitest';

import { parseDrawingBodyProperties } from '../../../drawingml/read-text-body';
import type { DrawingBodyProperties } from '../../../drawingml/text-model';
import { NS, parseXml } from '../../../xml/index';
import { PptxRuntimeDependencyFactory } from '../../core/core/factories/PptxRuntimeDependencyFactory';
import { PptxHandlerRuntime } from '../../core/core/runtime/PptxHandlerRuntimeImplementation';
import type { TextStyle, XmlObject } from '../../core/types';
import { objectBodyProperties } from '../../core/utils/body-properties-neutral';
import { canonicalJson, walkParityCorpus, type ParityVisit } from './drawingml-text-parity-corpus';

/**
 * DrawingML body property parity (docs/agnostic-core-plan.md step 3): the pptx reader and writer
 * of the `a:bodyPr` attributes and autofit choice run on the shared body codec
 * (`drawingml/text-body-codec`), and this harness proves the two paths agree over every
 * `a:bodyPr` of every part of every committed deck (`pptx/__tests__/fixtures`, including the
 * `e2e` snapshot).
 *
 * - Reading: each element is read through the pptx object tree (`objectBodyProperties`, what
 *   `applyBodyProperties` runs) and through the DOM (`parseDrawingBodyProperties`), paired by
 *   element path; the codec families must be deeply equal.
 * - Saving: the element is read into a `TextStyle` by the runtime (`applyBodyProperties`), written
 *   into a new text body by `applyTextBodyContent`, serialised and read through the DOM reader; the
 *   codec families must equal the source's once normalised to what the pptx model keeps (see
 *   `savedFamilies`: one column, a zero column gap, unknown overflow and wrap tokens and the
 *   anchor and direction spellings PowerPoint's model does not distinguish).
 *
 * Today: 5,158 parts, 26,999 elements and 53,047 codec fields read and saved; no differences.
 *
 * Before the switch, the previous pptx reader and writer were compared with the new ones over the
 * same corpus (26,999 `a:bodyPr` elements): every read `TextStyle` and every written `a:bodyPr`
 * was equal; 427 written nodes list their attributes in a different (schema) order.
 */

const BODY_ELEMENTS = new Set(['bodyPr']);
const builder = new PptxRuntimeDependencyFactory().createBuilder();
const runtime = new PptxHandlerRuntime() as unknown as {
	applyBodyProperties(bodyPr: XmlObject, txBody: undefined, style: TextStyle): unknown;
	applyTextBodyContent(
		shape: XmlObject,
		el: XmlObject,
		resolve: (target: string) => string | undefined,
		map: () => Map<string, string>,
	): void;
};
const canonical = canonicalJson;

/** The codec families of a DOM reading (no warp, 3D or kept children). */
function codecFamilies(value: DrawingBodyProperties | undefined): DrawingBodyProperties {
	const {
		extraXml: _extra,
		scene3d: _scene,
		shape3d: _shape,
		flatText3d: _flat,
		...rest
	} = value ?? {};
	return rest;
}

/** What the pptx body writer saves of an `a:bodyPr`, read back on the DOM. */
function savedThroughPptx(node: XmlObject): DrawingBodyProperties {
	const style: TextStyle = {};
	runtime.applyBodyProperties(node, undefined, style);
	const shape: XmlObject = {};
	runtime.applyTextBodyContent(
		shape,
		{ type: 'text', text: 'x', textStyle: style },
		() => undefined,
		() => new Map(),
	);
	const bodyPr = (shape['p:txBody'] as XmlObject)['a:bodyPr'] as XmlObject;
	const xml = builder.build({ 'a:bodyPr': { '@_xmlns:a': NS.a, ...bodyPr } }) as string;
	return codecFamilies(parseDrawingBodyProperties(parseXml(xml).documentElement));
}

const ANCHORS: Record<string, string> = { t: 't', top: 't', ctr: 'ctr', center: 'ctr', b: 'b' };
const OVERFLOWS = new Set(['overflow', 'clip', 'ellipsis']);

/** The source families as the pptx model keeps and writes them. */
function savedFamilies(value: DrawingBodyProperties): DrawingBodyProperties {
	const out = { ...value };
	if (out.columnCount !== undefined && out.columnCount <= 1) delete out.columnCount;
	if (out.columnSpacingEmu !== undefined && out.columnSpacingEmu <= 0) delete out.columnSpacingEmu;
	if (out.horizontalOverflow && !['overflow', 'clip'].includes(out.horizontalOverflow))
		delete out.horizontalOverflow;
	if (out.verticalOverflow && !OVERFLOWS.has(out.verticalOverflow)) delete out.verticalOverflow;
	if (out.wrap && !['none', 'square'].includes(out.wrap.toLowerCase())) delete out.wrap;
	if (out.anchor) {
		const anchor = ANCHORS[out.anchor.toLowerCase()];
		if (anchor) out.anchor = anchor;
	}
	if (out.vertical?.toLowerCase() === 'horz') delete out.vertical;
	// The model keeps no `textNoShape` warp (it is "no warp"), and keeps `adj`/`adj1`/`adj2` only.
	if (out.textWarp?.preset === 'textNoShape') delete out.textWarp;
	else if (out.textWarp) {
		const guides = (out.textWarp.guides ?? []).filter(
			(guide) =>
				['adj', 'adj1', 'adj2'].includes(guide.name) && /^val\s+-?\d+$/u.test(guide.formula.trim()),
		);
		out.textWarp = {
			preset: out.textWarp.preset,
			guides: guides.map((guide) => ({ name: guide.name, formula: guide.formula.trim() })),
		};
	}
	if (out.autofit?.kind === 'normal') {
		const { fontScale, lineSpacingReduction } = out.autofit;
		out.autofit = {
			kind: 'normal',
			...(fontScale !== undefined && fontScale > 0 && fontScale < 1 ? { fontScale } : {}),
			...(lineSpacingReduction !== undefined && lineSpacingReduction > 0
				? { lineSpacingReduction }
				: {}),
		};
	}
	return out;
}

interface Difference {
	label: string;
	objectTree: string;
	dom: string;
}

const readDifferences: Difference[] = [];
const saveDifferences: Difference[] = [];
let elementsCompared = 0;
let fieldsCompared = 0;
let partsCompared = 0;

function compareBody({ label, element, node }: ParityVisit): void {
	const dom = codecFamilies(parseDrawingBodyProperties(element));
	const objectTree = objectBodyProperties(node);
	elementsCompared++;
	fieldsCompared += Object.keys(dom).length;
	if (canonical(objectTree) !== canonical(dom))
		readDifferences.push({ label, objectTree: canonical(objectTree), dom: canonical(dom) });
	const saved = savedThroughPptx(node);
	if (canonical(saved) !== canonical(savedFamilies(dom)))
		saveDifferences.push({
			label,
			objectTree: canonical(saved),
			dom: canonical(savedFamilies(dom)),
		});
}

beforeAll(async () => {
	partsCompared = await walkParityCorpus(/bodyPr\b/u, BODY_ELEMENTS, compareBody);
}, 600_000);

describe('DrawingML body properties: pptx object tree vs shared DOM reader', () => {
	it('compares a real corpus', () => {
		expect(partsCompared).toBeGreaterThan(5000);
		expect(elementsCompared).toBeGreaterThan(26_000);
		expect(fieldsCompared).toBeGreaterThan(50_000);
	});

	it('reads every body properties element the same way', () => {
		expect(readDifferences.slice(0, 5)).toEqual([]);
		expect(readDifferences).toHaveLength(0);
	});

	it('saves every body properties element back to the same codec values', () => {
		expect(saveDifferences.slice(0, 5)).toEqual([]);
		expect(saveDifferences).toHaveLength(0);
	});
});
