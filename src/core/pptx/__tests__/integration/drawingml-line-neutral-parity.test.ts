import { beforeAll, describe, expect, it } from 'vitest';

import { readDrawingLineFields, type DrawingLineFields } from '../../../drawingml/line-read';
import { DOM_DRAWING_VIEW } from '../../../drawingml/node-view';
import { OBJECT_DRAWING_VIEW } from '../../core/color/object-drawing-view';
import { canonicalJson, walkParityCorpus, type ParityVisit } from './drawingml-text-parity-corpus';

/**
 * DrawingML line parity (docs/agnostic-core-plan.md step 3): the shared `a:ln` field reader
 * (`drawingml/line-read.ts`) serves the DOM line reader (`parseDrawingLine`) and the pptx text
 * outline and underline line readers (`extractTextRunStyle`) over the object-tree view. This
 * harness reads every `a:ln` and `a:uLn` of every committed deck through both views, paired by
 * element path, and requires the same width, dash, cap, compound, alignment and custom dash.
 *
 * Joins and ends are compared only when they carry attributes: the object view does not return an
 * empty element (`<a:round/>`, `<a:headEnd/>`) as a child, by design (see `object-drawing-view`),
 * and the pptx text readers do not read joins or ends through it.
 *
 * Before the switch, the previous pptx text outline and underline readers were compared with the
 * new ones through the runtime (`extractTextRunStyle` over 82,447 run property elements): identical.
 */

const LINE_ELEMENTS = new Set(['ln', 'uLn']);

function compared(fields: DrawingLineFields): DrawingLineFields {
	const { join: _join, headEnd, tailEnd, customDash, ...rest } = fields;
	return {
		...rest,
		...(customDash ? { customDash: customDash.map(({ dash, space }) => ({ dash, space })) } : {}),
		...(headEnd && Object.keys(headEnd).length ? { headEnd } : {}),
		...(tailEnd && Object.keys(tailEnd).length ? { tailEnd } : {}),
	};
}

const differences: { label: string; objectTree: string; dom: string }[] = [];
let elementsCompared = 0;
let partsCompared = 0;

function compareLine({ label, element, node }: ParityVisit): void {
	if (!node || Object.keys(node).length === 0) return;
	elementsCompared++;
	const dom = canonicalJson(compared(readDrawingLineFields(DOM_DRAWING_VIEW, element)));
	const objectTree = canonicalJson(compared(readDrawingLineFields(OBJECT_DRAWING_VIEW, node)));
	if (dom !== objectTree) differences.push({ label, objectTree, dom });
}

beforeAll(async () => {
	partsCompared = await walkParityCorpus(/<a:u?[Ll]n\b/u, LINE_ELEMENTS, compareLine);
}, 600_000);

describe('DrawingML lines: pptx object tree vs shared DOM reader', () => {
	it('compares a real corpus', () => {
		expect(partsCompared).toBeGreaterThan(1200);
		expect(elementsCompared).toBeGreaterThan(7500);
	});

	it('reads every line the same way', () => {
		expect(differences.slice(0, 5)).toEqual([]);
		expect(differences).toHaveLength(0);
	});
});
