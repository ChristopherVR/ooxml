import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { generateUnmatchedFadeInAnimations } from './morph-animation';
import { matchMorphElementsFull } from './morph-matching';
import { coveredStackedCopyIds } from './morph-stacked-copies';
import { buildMorphTransitionPlan } from './morph-plan';
import { morphOffstageTextSizeChange } from './morph-text-slot';

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

describe('stacked outgoing copies', () => {
	it('drops the covered copy as the dissolve starts, while the top copy fades', async () => {
		const { generateUnmatchedFadeOutAnimations } = await import('./morph-animation');
		const lower = text('lower', 'L', { x: 428, y: 273, width: 46, height: 87 });
		const upper = text('upper', 'L', { x: 426, y: 274, width: 46, height: 87 });
		const [dropped, fading] = generateUnmatchedFadeOutAnimations([lower, upper], 2000, 0);
		expect(dropped?.keyframes).toMatch(/%, 100% \{\s*opacity: 0;/u);
		expect(fading?.keyframes).not.toMatch(/%, 100% \{/u);
	});
});

describe('offstage text with a changed run size', () => {
	const from = {
		...text('from', 'Heading\nBody', { x: -500, width: 200 }),
		textSegments: [
			{ text: 'Heading', style: { fontSize: 14, fontFamily: 'Arial' } },
			{ text: '\n' },
			{ text: 'Body', style: { fontSize: 18, fontFamily: 'Arial' } },
		],
	} as PptxElement;
	const to = {
		...text('to', 'Heading\nBody', { x: 300, width: 200, height: 210 }),
		textSegments: [
			{ text: 'Heading', style: { fontSize: 18, fontFamily: 'Arial' } },
			{ text: '\n' },
			{ text: 'Body', style: { fontSize: 18, fontFamily: 'Arial' } },
		],
	} as PptxElement;

	it('uses one layout when the source starts wholly off the slide', () => {
		expect(morphOffstageTextSizeChange(from, to)).toBe(true);
		const plan = buildMorphTransitionPlan(slide([from]), slide([to]), 1000)!;
		expect(plan.outgoingElements.filter((element) => element.id === from.id)).toHaveLength(0);
		expect(plan.incomingAnimations.get(to.id)).not.toContain('-fade');
	});

	it('uses the destination layout on the ghost and hides the stage copy over a backdrop', () => {
		const backdrop = {
			id: 'backdrop',
			type: 'shape',
			x: -600,
			y: 0,
			width: 1200,
			height: 300,
		} as PptxElement;
		const plan = buildMorphTransitionPlan(slide([backdrop, from]), slide([to]), 1000)!;
		const ghost = plan.outgoingElements.find((element) => element.id === from.id);
		expect(ghost).toBeDefined();
		expect(plan.outgoingTextReplacements?.get(from.id)).toBe(to.id);
		expect(ghost?.height).toBe(210);
		expect(ghost && 'textSegments' in ghost && ghost.textSegments).toEqual(to.textSegments);
		const ghostName = plan.outgoingAnimations.get(from.id)?.split(' ')[0];
		const ghostBlock = plan.keyframesCss
			.split(`@keyframes ${ghostName} {`)[1]
			?.split('@keyframes')[0];
		expect(ghostBlock).toContain('translate(800px, 0px) rotate(0deg) scale(1, 1)');
		const name = plan.incomingAnimations.get(to.id)?.split(' ')[0];
		expect(name).toBeDefined();
		const block = plan.keyframesCss.split(`@keyframes ${name} {`)[1]?.split('@keyframes')[0];
		expect(block?.match(/opacity: 0;/gu)).toHaveLength(2);
	});

	it('does not suppress crossfades for visible or differently styled text', () => {
		expect(morphOffstageTextSizeChange({ ...from, x: 0 }, to)).toBe(false);
		expect(morphOffstageTextSizeChange(from, { ...to, text: 'Other' })).toBe(false);
		const changedFont = {
			...to,
			textSegments: [
				{ text: 'Heading', style: { fontSize: 18, fontFamily: 'Calibri' } },
				{ text: '\n' },
				{ text: 'Body', style: { fontSize: 18, fontFamily: 'Arial' } },
			],
		} as PptxElement;
		expect(morphOffstageTextSizeChange(from, changedFont)).toBe(false);
		expect(morphOffstageTextSizeChange(from, { ...to, shapeStyle: { fillColor: '#ff0000' } })).toBe(
			false,
		);
	});
});
