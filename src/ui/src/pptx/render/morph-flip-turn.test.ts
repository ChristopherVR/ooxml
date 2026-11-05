import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { buildMorphFlipTurn, morphFlipTurns } from './morph-flip-turn';

function triangle(extra: Record<string, unknown>): PptxElement {
	return {
		id: 't',
		type: 'shape',
		shapeType: 'rtTriangle',
		x: 133.333,
		y: 133.333,
		width: 400,
		height: 266.667,
		...extra,
	} as unknown as PptxElement;
}

type Vec = [number, number, number, number];

/** One CSS transform function applied to a homogeneous point. */
function apply(name: string, args: number[], [x, y, z, w]: Vec): Vec {
	const rad = ((args[0] ?? 0) * Math.PI) / 180;
	switch (name) {
		case 'translate':
			return [x + (args[0] ?? 0) * w, y + (args[1] ?? 0) * w, z, w];
		case 'perspective':
			return [x, y, z, w - z / (args[0] ?? 1)];
		case 'rotate':
			return [x * Math.cos(rad) - y * Math.sin(rad), x * Math.sin(rad) + y * Math.cos(rad), z, w];
		case 'rotateY':
			return [x * Math.cos(rad) + z * Math.sin(rad), y, -x * Math.sin(rad) + z * Math.cos(rad), w];
		case 'rotateX':
			return [x, y * Math.cos(rad) - z * Math.sin(rad), y * Math.sin(rad) + z * Math.cos(rad), w];
		case 'scale':
			return [x * (args[0] ?? 1), y * (args[1] ?? args[0] ?? 1), z, w];
		case 'scaleX':
			return [x * (args[0] ?? 1), y, z, w];
		case 'scaleY':
			return [x, y * (args[0] ?? 1), z, w];
		default:
			throw new Error(name);
	}
}

/** Where a CSS transform list puts local `(x, y)` (about the box centre), projected. */
function project(transform: string, x: number, y: number): [number, number] {
	const fns = [...transform.matchAll(/(\w+)\(([^)]*)\)/gu)].map(
		(m) => [m[1] ?? '', (m[2] ?? '').split(',').map((v) => Number.parseFloat(v))] as const,
	);
	let p: Vec = [x, y, 0, 1];
	for (const [name, args] of fns.reverse()) p = apply(name, args, p);
	return [p[0] / p[3], p[1] / p[3]];
}

describe('morph flip turn', () => {
	const from = triangle({});
	const to = triangle({ x: 666.667, y: 373.333, flipHorizontal: true });

	it('turns only a pair whose flip changes', () => {
		expect(morphFlipTurns(from, to)).toBe(true);
		expect(morphFlipTurns(from, triangle({ x: 500 }))).toBe(false);
	});

	it('looks at the turn from the slide centre, 60 degrees across', () => {
		// The eye sits (1280 / 2) / tan(30 deg) = 1108.5 px in front of the slide.
		const { start, end } = buildMorphFlipTurn(to, from, to, 0, 0);
		expect(start).toContain('perspective(1108.513px)');
		// The incoming half's frame: centre at (866.67, 506.67), so the slide
		// centre (640, 360) is (-226.67, -146.67) from it.
		expect(start).toContain('translate(-226.667px, -146.667px) perspective(');
		expect(start).toContain('rotateY(0deg)');
		expect(end).toContain('rotateY(-180deg)');
		// Both frames carry the outgoing (unflipped) mirror: the turn makes it.
		expect(start).toContain('scaleX(1) scaleY(1)');
		expect(end).toContain('scaleX(1) scaleY(1)');
	});

	it('ends on the flipped shape and starts on the outgoing one', () => {
		const { start, end } = buildMorphFlipTurn(to, from, to, 0, 0);
		// Local x = +200 (the right edge, unflipped) lands mirrored at the end...
		const [ex] = project(end, 200, 0);
		expect(ex).toBeCloseTo(-200, 3);
		// ...and on the outgoing box at the start (533.3 px to the left).
		const [sx, sy] = project(start, 200, 0);
		expect(sx).toBeCloseTo(-333.333, 2);
		expect(sy).toBeCloseTo(-240, 2);
	});

	it('swings the shape-local +x side towards the viewer, whichever way it flips', () => {
		// Midway through PowerPoint's turn the right edge of an unflipped triangle
		// is the near (taller) side; flipping back, its mirrored left is.
		const half = (transform: string) => transform.replace(/rotateY\([^)]*\)/u, 'rotateY(-90deg)');
		const { end } = buildMorphFlipTurn(to, from, to, 0, 0);
		const mid = half(end).replace('rotateY(-90deg)', 'rotateY(-60deg)');
		const near = Math.abs(project(mid, 200, 100)[1] - project(mid, 200, -100)[1]);
		const far = Math.abs(project(mid, -200, 100)[1] - project(mid, -200, -100)[1]);
		expect(near).toBeGreaterThan(far);

		const back = buildMorphFlipTurn(from, to, from, 0, 0);
		expect(back.end).toContain('rotateY(180deg)');
		expect(back.start).toContain('scaleX(-1)');
	});
});
