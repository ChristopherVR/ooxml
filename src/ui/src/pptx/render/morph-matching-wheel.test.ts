/**
 * Regression guard for issue #34, on the reporter's own wheel deck
 * (`solution-explorer.pptx`): clicking round the wheel's topics, every wedge
 * must morph in place (recolouring when the selection moves), never glide
 * into another sector.
 *
 * The identical-twin pass used to pair each wedge with the FIRST same-box,
 * same-fill wedge in document order. Wedges of one orientation share a box
 * and the unselected ones share a fill, so from the third topic on wedges
 * paired with their mirror across the wheel and visibly swapped places.
 */
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { PptxHandler } from 'ooxml-core/pptx';
import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { beforeAll, describe, expect, it } from 'vitest';

import { matchMorphElementsFull } from './morph-matching';

const FIXTURE = resolve(__dirname, '../../../../../e2e/pptx/fixtures/solution-explorer.pptx');

/** The topic slides, in the order the wheel's sectors run clockwise. */
const TOPICS = [2, 3, 4, 5, 6, 7, 8, 9];

const isWedge = (el: PptxElement): boolean => (el.name ?? '').startsWith('Free-form');

describe('morph matching on the issue #34 wheel deck', () => {
	let slides: PptxSlide[];

	beforeAll(async () => {
		const bytes = readFileSync(FIXTURE);
		const data = await new PptxHandler().load(
			bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
		);
		slides = data.slides;
	}, 60000);

	const steps = TOPICS.map((from, i) => [from, TOPICS[(i + 1) % TOPICS.length]] as const);
	it.each(steps)('keeps every wedge in its sector from topic %i to %i', (from, to) => {
		const { pairs } = matchMorphElementsFull(slides[from], slides[to]);
		const wedgePairs = pairs.filter((p) => isWedge(p.fromElement) && isWedge(p.toElement));
		expect(wedgePairs.length).toBeGreaterThanOrEqual(8);
		for (const { fromElement, toElement } of wedgePairs) {
			expect([toElement.x, toElement.y]).toEqual([fromElement.x, fromElement.y]);
		}
	});
});
