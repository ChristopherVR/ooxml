import type { TextSegment } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { buildParagraphRuns } from './paragraph-run-build';

function build(segments: TextSegment[]) {
	return buildParagraphRuns({
		paraSegments: segments,
		paraIndices: segments.map((_, index) => index),
		markerSegment: undefined,
		fontScale: 1,
		blockFont: { fontSizePx: 37 } as never,
		blockScriptStyle: undefined,
		tabStops: undefined,
		defaultTabSize: undefined,
		fontAlignment: undefined,
		fieldContext: undefined,
	});
}

const text = (value: string): TextSegment => ({ text: value, style: { fontSize: 24 } });
const softBreak = (): TextSegment => ({ text: '\n', isLineBreak: true, style: { fontSize: 24 } });

describe('issue 249: a paragraph ending in a:br keeps its empty line', () => {
	it('adds a zero-width run at the break size after a trailing break', () => {
		const runs = build([text('Lorem ipsum'), softBreak()]);
		const last = runs[runs.length - 1]!;
		expect(runs[runs.length - 2]!.text).toBe('\n');
		expect(last.text).toBe('​');
		expect(last.style.fontSize).toBe(runs[runs.length - 2]!.style.fontSize);
		expect(last.style.fontSize).toBeDefined();
		expect(last.segmentIndex).toBe(runs[runs.length - 2]!.segmentIndex);
		expect(last.segmentIndex).toBe(1);
	});

	it('leaves a break between two runs alone', () => {
		const runs = build([text('one'), softBreak(), text('two')]);
		expect(runs.some((run) => run.text === '​')).toBe(false);
	});

	it('leaves a paragraph without a break alone', () => {
		const runs = build([text('one')]);
		expect(runs.some((run) => run.text === '​')).toBe(false);
	});
});
