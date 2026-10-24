import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { generateUnmatchedFadeInAnimations } from './morph-animation';
import { matchMorphElementsFull } from './morph-matching';
import { coveredStackedCopyIds } from './morph-stacked-copies';

function text(
	id: string,
	body: string,
	overrides: Partial<PptxElement> & { creationId?: string } = {},
): PptxElement {
	const { creationId, ...rest } = overrides;
	return {
		id,
		type: 'text',
		x: 0,
		y: 0,
		width: 300,
		height: 200,
		text: body,
		...(creationId
			? {
					rawXml: {
						'p:nvSpPr': {
							'p:cNvPr': {
								'a:extLst': { 'a:ext': { 'a16:creationId': { '@_id': creationId } } },
							},
						},
					},
				}
			: {}),
		...rest,
	} as PptxElement;
}

const slide = (elements: PptxElement[]): PptxSlide => ({ id: 's', elements }) as PptxSlide;

const ids = (pairs: { fromElement: PptxElement; toElement: PptxElement }[]): string[] =>
	pairs.map((pair) => `${pair.fromElement.id}>${pair.toElement.id}`);

// Issue #141, slides 17 -> 18: two columns whose paragraphs are reshuffled
// between them, and a third column parked off-stage that lands under a new name.
describe('morph text flow (issue #141, slides 17 -> 18)', () => {
	const left17 = text('l17', 'Heading\nShared first paragraph\nMoves to the right column', {
		shapeId: '20',
		name: 'Object 20',
	});
	const right17 = text('r17', 'Right column opener\nRight column only', {
		x: 700,
		creationId: '{R}',
		name: 'Object 7',
	});
	const left18 = text('l18', 'Heading\nShared first paragraph\nA new paragraph', {
		x: -500,
		shapeId: '20',
		name: 'Object 20',
	});
	const right18 = text('r18', 'Right column opener\nA rewrite\nMoves to the right column', {
		x: 1300,
		creationId: '{R}',
		name: 'Object 7',
	});
	const parked = text('p17', 'Heading\nThird column words\nMore third column words', {
		x: -600,
		name: 'Object 3',
	});
	const landed = text('p18', 'Heading\nThird column words\nMore third column words', {
		x: 870,
		height: 205,
		name: 'Object 6',
	});

	it('does not carry a box across when its paragraphs land in different boxes', () => {
		const result = matchMorphElementsFull(
			slide([left17, right17, parked]),
			slide([left18, right18, landed]),
		);
		expect(ids(result.pairs)).not.toContain('l17>l18');
		expect(ids(result.pairs)).not.toContain('r17>r18');
		expect(result.unmatchedFrom.map((el) => el.id)).toEqual(['l17', 'r17']);
	});

	it('glides a box that says the same thing under another name', () => {
		const result = matchMorphElementsFull(
			slide([left17, right17, parked]),
			slide([left18, right18, landed]),
		);
		expect(ids(result.pairs)).toContain('p17>p18');
	});

	it('still carries a box whose paragraphs all stay in it', () => {
		const edited = text('l18', 'Heading\nShared first paragraph\nA new paragraph', {
			x: -500,
			shapeId: '20',
			name: 'Object 20',
		});
		const result = matchMorphElementsFull(slide([left17]), slide([edited]));
		expect(ids(result.pairs)).toEqual(['l17>l18']);
	});

	it('leaves repeated text to the other passes', () => {
		const a = text('a', 'Same words here', { x: -600 });
		const b = text('b', 'Same words here', { x: -600, y: 300 });
		const c = text('c', 'Same words here', { x: 900 });
		const result = matchMorphElementsFull(slide([a, b]), slide([c]));
		expect(result.pairs).toHaveLength(0);
	});
});

// Issue #141, slides 6 -> 7: two "L" text boxes stacked 2px apart fade in.
describe('stacked incoming copies', () => {
	const lower = text('lower', 'L', { x: 428, y: 273, width: 46, height: 87 });
	const upper = text('upper', 'L', { x: 426, y: 274, width: 46, height: 87 });

	it('finds the copy covered by an almost identical one above it', () => {
		expect([...coveredStackedCopyIds([lower, upper])]).toEqual(['lower']);
		expect(coveredStackedCopyIds([lower, text('far', 'L', { x: 500, y: 273 })]).size).toBe(0);
		expect(coveredStackedCopyIds([lower, text('other', 'M', { x: 426, y: 274 })]).size).toBe(0);
	});

	it('holds the covered copy hidden so only one glyph dissolves in', () => {
		const [hidden, fading] = generateUnmatchedFadeInAnimations([lower, upper], 2000, 0);
		expect(hidden?.keyframes).toMatch(/0%, 99\.9% \{\s*opacity: 0;/u);
		expect(hidden?.keyframes).toMatch(/100% \{\s*opacity: 1;/u);
		expect(fading?.keyframes).not.toContain('99.9%');
	});
});
