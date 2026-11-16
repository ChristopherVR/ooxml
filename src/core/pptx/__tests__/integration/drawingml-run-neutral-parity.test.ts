import { beforeAll, describe, expect, it } from 'vitest';

import { parseDrawingRunProperties } from '../../../drawingml/read-run-properties';
import type { DrawingRunProperties } from '../../../drawingml/text-model';
import { NS, parseXml } from '../../../xml/index';
import { PptxRuntimeDependencyFactory } from '../../core/core/factories/PptxRuntimeDependencyFactory';
import { PptxHandlerRuntime } from '../../core/core/runtime/PptxHandlerRuntimeImplementation';
import type { TextStyle, XmlObject } from '../../core/types';
import { objectRunProperties } from '../../core/utils/run-properties-neutral';
import { canonicalJson, walkParityCorpus, type ParityVisit } from './drawingml-text-parity-corpus';

/**
 * DrawingML run property parity (docs/agnostic-core-plan.md step 3): the pptx readers and writer
 * of the `a:rPr`, `a:defRPr` and `a:endParaRPr` attributes, font slots, run `a:rtl` and underline
 * follow-text markers run on the shared run codec (`drawingml/text-run-codec`), and this harness
 * proves the two paths agree over every such element of every part of every committed deck
 * (`pptx/__tests__/fixtures`, including the `e2e` snapshot).
 *
 * - Reading: each element is read through the pptx object tree (`objectRunProperties`, what
 *   `extractTextRunStyle`, the SmartArt projection parser and the level-style reader run) and
 *   through the DOM (`parseDrawingRunProperties`), paired by element path; the codec families must
 *   be deeply equal.
 * - Saving: the element is read into a `TextStyle` by the runtime (`extractTextRunStyle`), written
 *   back by `createRunPropertiesFromTextStyle`, serialised and read through the DOM reader; the
 *   codec families must equal the source's, except the defaults that writer adds or drops for a
 *   style without an inherited baseline (`lang="en-US"` and `dirty="0"` when absent; `spc`, `kern`
 *   and `baseline` of 0), which the slide save resolves against the paragraph baseline.
 *
 * Today: 4,810 parts, 80,232 elements and 150,587 codec fields read and saved; no differences.
 *
 * Before the switch, the previous pptx readers and writer were compared with the new ones over
 * the same corpus (80,319 run property elements and 35,226 level elements): every runtime run and
 * default style, every SmartArt-projection reading, every level style and every written `a:rPr`
 * was equal. Only the attribute order of a rebuilt `a:rPr` changed (2,311 elements), to schema
 * order; 1,566 of those now match the source file's own order, which the old order never did.
 */

const RUN_ELEMENTS = new Set(['rPr', 'defRPr', 'endParaRPr']);
const builder = new PptxRuntimeDependencyFactory().createBuilder();
const runtime = new PptxHandlerRuntime() as unknown as {
	extractTextRunStyle(node: XmlObject, align: TextStyle['align']): TextStyle;
	createRunPropertiesFromTextStyle(style: TextStyle): XmlObject | undefined;
};
const canonical = canonicalJson;

/** The codec families of a DOM reading (no fills, outline, highlight, links or kept children). */
function codecFamilies(value: DrawingRunProperties | undefined): DrawingRunProperties {
	const {
		extraXml: _extra,
		line: _line,
		lineSourceXml: _lineXml,
		fill: _fill,
		highlight: _highlight,
		underlineLineXml: _uLn,
		underlineFill: _uFill,
		hyperlinkClick: _click,
		hyperlinkMouseOver: _over,
		...rest
	} = value ?? {};
	return rest;
}

/** What the pptx run writer saves of a run properties node, read back on the DOM. */
function savedThroughPptx(node: XmlObject): DrawingRunProperties {
	const rPr = runtime.createRunPropertiesFromTextStyle(runtime.extractTextRunStyle(node, 'left'));
	const xml = builder.build({ 'a:rPr': { '@_xmlns:a': NS.a, ...rPr } }) as string;
	return codecFamilies(parseDrawingRunProperties(parseXml(xml).documentElement));
}

/** The source families with the baseline-free writer's defaults applied. */
function savedFamilies(value: DrawingRunProperties): DrawingRunProperties {
	const out: DrawingRunProperties = { language: 'en-US', dirty: false, ...value };
	if (out.spacingPt === 0) delete out.spacingPt;
	if (out.kernPt === 0) delete out.kernPt;
	if (out.baseline === 0) delete out.baseline;
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

function compareRun({ label, element, node }: ParityVisit): void {
	const dom = codecFamilies(parseDrawingRunProperties(element));
	const objectTree = objectRunProperties(node);
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
	partsCompared = await walkParityCorpus(/RPr\b/u, RUN_ELEMENTS, compareRun);
}, 600_000);

describe('DrawingML run properties: pptx object tree vs shared DOM reader', () => {
	it('compares a real corpus', () => {
		expect(partsCompared).toBeGreaterThan(4500);
		expect(elementsCompared).toBeGreaterThan(80_000);
		expect(fieldsCompared).toBeGreaterThan(150_000);
	});

	it('reads every run properties element the same way', () => {
		expect(readDifferences.slice(0, 5)).toEqual([]);
		expect(readDifferences).toHaveLength(0);
	});

	it('saves every run properties element back to the same codec values', () => {
		expect(saveDifferences.slice(0, 5)).toEqual([]);
		expect(saveDifferences).toHaveLength(0);
	});
});
