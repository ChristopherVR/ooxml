import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../../index';

// Issue 249: PowerPoint ends each paragraph with `<a:r/>...<a:br/>`. The runs and the break
// are already grouped by tag, and the legacy collapsed-break repair used to spread that one
// trailing break into a phantom line break after the first run of every paragraph.
describe('issue 249: trailing a:br after grouped runs', () => {
	it('keeps one break per authored a:br, at the end of its paragraph', async () => {
		const bytes = readFileSync(
			new URL('../../../__tests__/fixtures/issue-249-random-newlines.pptx', import.meta.url),
		);
		const document = await new PptxHandler().load(
			bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
		);
		const element = document.slides[0]!.elements[0] as {
			textSegments: Array<{ text: string; isLineBreak?: boolean }>;
		};
		const segments = element.textSegments;
		expect(segments.filter((segment) => segment.isLineBreak)).toHaveLength(4);
		segments.forEach((segment, index) => {
			if (!segment.isLineBreak) return;
			// A break closes its paragraph: the next segment is the paragraph terminator or nothing.
			const next = segments[index + 1];
			expect(next === undefined || (next.text === '\n' && !next.isLineBreak)).toBe(true);
		});
	});
});
