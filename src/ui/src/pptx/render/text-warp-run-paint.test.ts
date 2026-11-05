import { describe, expect, it } from 'vitest';

import { resolveWarpRunPaint } from './text-warp-run-paint';

describe('resolveWarpRunPaint', () => {
	it('fills with the run colour, falling back to the element then black', () => {
		expect(
			resolveWarpRunPaint({ text: 'a', style: { color: '#112233' } }, { color: '#445566' }).fill,
		).toBe('#112233');
		expect(resolveWarpRunPaint({ text: 'a' }, { color: '#445566' }).fill).toBe('#445566');
		expect(resolveWarpRunPaint({ text: 'a' }, {}).fill).toBe('#000000');
	});

	it('resolves weight, style and decoration', () => {
		const paint = resolveWarpRunPaint(
			{ text: 'a', style: { bold: true, italic: true, underline: true, fontSize: 30 } },
			{},
		);
		expect(paint.fontWeight).toBe('bold');
		expect(paint.fontStyle).toBe('italic');
		expect(paint.textDecoration).toBe('underline');
		expect(paint.fontSizePx).toBe(30);
	});

	it('uses the caller default size and colour only when nothing declares one', () => {
		expect(resolveWarpRunPaint({ text: 'a' }, {}, '#111827', 24)).toMatchObject({
			fill: '#111827',
			fontSizePx: 24,
		});
		expect(resolveWarpRunPaint({ text: 'a' }, { fontSize: 40 }, '#111827', 24).fontSizePx).toBe(40);
	});
});
