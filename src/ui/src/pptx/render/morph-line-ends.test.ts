import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { matchMorphElementsFull } from './morph-matching';
import { morphArrowSignature, morphLineEndsBreakPair } from './morph-line-ends';
import { buildMorphTransitionPlan } from './morph-plan';

const slide = (id: string, elements: PptxElement[]): PptxSlide =>
	({ id, slideNumber: 1, elements }) as unknown as PptxSlide;

const line = (id: string, style: Record<string, unknown> = {}): PptxElement =>
	({
		id,
		name: 'Line',
		type: 'connector',
		shapeType: 'line',
		x: 100,
		y: 100,
		width: 600,
		height: 0,
		shapeStyle: { strokeColor: '#000000', strokeWidth: 4, ...style },
	}) as unknown as PptxElement;

const pairs = (a: PptxElement, b: PptxElement) =>
	matchMorphElementsFull(slide('a', [a]), slide('b', [b])).pairs.length;

describe('morph line ends', () => {
	it('does not pair a line that gains or loses an arrowhead', () => {
		const plain = line('a-1');
		const arrow = line('b-1', { connectorEndArrow: 'triangle' });
		expect(morphLineEndsBreakPair(plain, arrow)).toBe(true);
		expect(pairs(plain, arrow)).toBe(0);
		expect(pairs(arrow, plain)).toBe(0);
	});

	it('does not pair a dashed line with a solid one', () => {
		const dashed = line('a-1', { strokeDash: 'dash' });
		const solid = line('b-1', { strokeDash: 'solid' });
		expect(morphLineEndsBreakPair(dashed, solid)).toBe(true);
		expect(pairs(dashed, solid)).toBe(0);
	});

	it('pairs a line whose head only changes type or size, and dissolves the heads', () => {
		const a = line('a-1', { connectorEndArrow: 'triangle' });
		const b = line('b-1', { connectorEndArrow: 'oval' });
		const big = line('c-1', {
			connectorEndArrow: 'triangle',
			connectorEndArrowWidth: 'lg',
			connectorEndArrowLength: 'lg',
		});
		expect(pairs(a, b)).toBe(1);
		expect(pairs(a, big)).toBe(1);
		expect(morphArrowSignature({ connectorEndArrow: 'triangle' })).toBe('>triangle/med/med');
		expect(morphArrowSignature({ connectorEndArrow: 'none' })).toBe('>');
		const plan = buildMorphTransitionPlan(slide('a', [a]), slide('b', [big]), 2000)!;
		expect(plan.crossfadeGroups).toHaveLength(1);
	});

	it('treats a default head size and an explicit medium one as the same line', () => {
		const implicit = line('a-1', { connectorEndArrow: 'triangle' });
		const explicit = line('b-1', {
			connectorEndArrow: 'triangle',
			connectorEndArrowWidth: 'med',
			connectorEndArrowLength: 'med',
		});
		const plan = buildMorphTransitionPlan(slide('a', [implicit]), slide('b', [explicit]), 2000);
		expect(plan?.crossfadeGroups ?? []).toHaveLength(0);
	});

	it('keeps a !! named pair together through an arrowhead change', () => {
		const a = { ...line('a-1'), name: '!!Rule' } as PptxElement;
		const b = { ...line('b-1', { connectorEndArrow: 'triangle' }), name: '!!Rule' } as PptxElement;
		expect(pairs(a, b)).toBe(1);
	});
});
