import { describe, expect, it } from 'vitest';
import type { DrawingAnchor } from '../model';
import { slicerFrameXml } from './frame';

const anchor: DrawingAnchor = {
	from: { col: 1, row: 1, colOffset: 0, rowOffset: 0 },
	to: { col: 4, row: 10, colOffset: 0, rowOffset: 0 },
};
const box = { x: 0, y: 0, cx: 100, cy: 100 };
const ids = (xml: string) => [...xml.matchAll(/<xdr:cNvPr id="(\d+)" name="([^"]*)"/g)];

describe('slicer frame fallback', () => {
	it.each(['slicer', 'timeline'] as const)(
		'gives the %s fallback shape the id and name of its frame',
		(kind) => {
			const first = slicerFrameXml(kind, 'Region & "x"', 3, anchor, box);
			const second = slicerFrameXml(kind, 'Year', 4, anchor, box);
			expect(first).not.toContain('id="0"');
			expect(ids(first).map((match) => [match[1], match[2]])).toEqual([
				['3', 'Region &amp; &quot;x&quot;'],
				['3', 'Region &amp; &quot;x&quot;'],
			]);
			expect(ids(second).map((match) => match[1])).toEqual(['4', '4']);
		},
	);
});
