import { expect, it } from 'vitest';
import { demoDocument } from './demo-document';
import { visioStraightLineHandles } from './line-edit';

const line = () => ({
	...structuredClone(demoDocument.pages[0]!.shapes[0]!),
	kind: 'connector' as const,
	width: 2,
	height: 0,
	geometry: [{ path: 'M 0 0 L 2 0', fill: false, stroke: true }],
});
it('admits visible local straight-line handles with output rounding', () => {
	const shape = line();
	shape.width = 2.0000004;
	expect(visioStraightLineHandles(shape)).toBe(true);
});
it.each([
	{ hidden: true },
	{ height: 1 },
	{ width: 0 },
	{ masterId: '1' },
	{ geometry: [{ path: 'M 0 0 L 1 1 L 2 0', fill: false, stroke: true }] },
	{ geometry: [{ path: 'M 0 0 L 1 0', fill: false, stroke: true }] },
])('excludes visually unsuitable handles (%s)', (override) => {
	expect(visioStraightLineHandles({ ...line(), ...override })).toBe(false);
});
