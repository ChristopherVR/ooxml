import { describe, expect, it } from 'vitest';
import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { matchMorphElementsFull } from './morph-matching';

const slide = (elements: PptxElement[]) => ({ id: 'slide', elements }) as PptxSlide;
const picture = (id: string, fields: Record<string, unknown>): PptxElement =>
	({
		id,
		type: 'picture',
		x: 302,
		y: 22,
		width: 675,
		height: 675,
		...fields,
	}) as PptxElement;

describe('unrelated artwork in the solution explorer morph (#105)', () => {
	it('does not rotate an incoming animation from a nearby wheel graphic', () => {
		const wheel = picture('wheel', { name: 'Graphic 23', svgPath: 'wheel.svg', rotation: 45 });
		const animation = picture('animation', {
			name: '!!Break',
			imagePath: 'radar.gif',
			x: 349,
			y: 67,
			width: 706,
			height: 407,
		});
		for (const [from, to] of [
			[wheel, animation],
			[animation, wheel],
		]) {
			const result = matchMorphElementsFull(slide([from!]), slide([to!]));
			expect(result.pairs).toEqual([]);
			expect(result.unmatchedTo).toEqual([to]);
		}
	});
	it('distinguishes different SVG artwork in the identical-paint pass', () => {
		const result = matchMorphElementsFull(
			slide([picture('wheel', { svgPath: 'wheel.svg', shapeStyle: { fillColor: '#fff' } })]),
			slide([picture('aircraft', { svgPath: 'aircraft.svg', shapeStyle: { fillColor: '#fff' } })]),
		);
		expect(result.pairs).toEqual([]);
	});
	it('still morphs replacement artwork with the same explicit identity', () => {
		const result = matchMorphElementsFull(
			slide([picture('before', { name: '!!scene', imagePath: 'before.png', rotation: 45 })]),
			slide([picture('after', { name: '!!scene', imagePath: 'after.png' })]),
		);
		expect(result.pairs).toHaveLength(1);
	});
});
