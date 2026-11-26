import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { inkStrokeFromSamples } from './edit';
import { inkStrokesEqual, readInkPart, updatedInkPart, writeInkPart } from './part';
import { inkSvg } from './svg';
import { inkObjectFromXml, inkObjectToXml } from './xml-tree';

/** The ink part of Word 16's own save (`docx/__fixtures__/ink`), difference-encoded traces. */
async function wordInkPart(): Promise<string> {
	const zip = await JSZip.loadAsync(
		readFileSync(
			path.join(import.meta.dirname, '../../docx/__fixtures__/ink/word-ink-inline.docx'),
		),
	);
	return zip.file('word/ink/ink1.xml')!.async('string');
}

const pen = { kind: 'pen' as const, color: '#0078D4', width: 2 };

describe('InkML parts', () => {
	it('reads Word’s part into strokes stretched onto the drawing box', async () => {
		const { strokes, raw } = readInkPart(await wordInkPart(), { width: 227, height: 76 });
		expect(strokes).toHaveLength(4);
		expect(strokes.map((stroke) => stroke.color)).toEqual([
			'#E81123',
			'#0078D7',
			'#0078D7',
			'#0078D7',
		]);
		// 0.05 cm and 0.1 cm brushes in pixels.
		expect(strokes[0]?.width).toBeCloseTo(1.89, 2);
		expect(strokes[1]?.width).toBeCloseTo(3.78, 2);
		const xs = strokes.flatMap((stroke) =>
			[...stroke.path.matchAll(/[ML] ([\d.-]+)/g)].map((m) => Number(m[1])),
		);
		expect(Math.min(...xs)).toBeCloseTo(0, 0);
		expect(Math.max(...xs)).toBeCloseTo(227, 0);
		expect(raw).toBeDefined();
	});

	it('keeps an untouched part and rewrites a changed one with its unknown nodes', async () => {
		const xml = await wordInkPart();
		const box = { width: 227, height: 76 };
		const { strokes } = readInkPart(xml, box);
		expect(updatedInkPart(xml, strokes, box)).toBeUndefined();
		const rewritten = updatedInkPart(xml, strokes.slice(1), box)!;
		expect(rewritten).toContain('<ink:ink');
		// The source's context (its timestamp, channel properties) is gone with the old definitions,
		// but nothing outside them is lost and the rewrite reads back as the kept strokes.
		const reread = readInkPart(rewritten, box).strokes;
		expect(inkStrokesEqual(reread, strokes.slice(1))).toBe(true);
	});

	it('writes new strokes that read back exactly, brushes in centimetres', () => {
		const stroke = inkStrokeFromSamples(
			[
				{ x: 1, y: 2, time: 10 },
				{ x: 30, y: 12, time: 30 },
				{ x: 60, y: 4, time: 55 },
			],
			pen,
		)!;
		const xml = writeInkPart([stroke]);
		expect(xml).toContain('name="width" value="0.05291666666666667" units="cm"');
		expect(xml).toContain('name="height"');
		const [back] = readInkPart(xml, { width: 61, height: 13 }).strokes;
		expect(back?.path).toBe(stroke.path);
		expect(back?.width).toBe(2);
		expect(back?.color).toBe('#0078D4');
		expect(back?.durationMs).toBe(45);
	});

	it('reads nothing from a malformed or oversized part instead of throwing', () => {
		expect(readInkPart('<ink:ink').strokes).toEqual([]);
		expect(readInkPart(undefined).strokes).toEqual([]);
		expect(readInkPart('x'.repeat(9 * 1024 * 1024)).strokes).toEqual([]);
	});

	it('round-trips the object tree, escaping text and attributes', () => {
		const xml =
			'<?xml version="1.0"?><a:root xmlns:a="urn:a" note="&quot;x&lt;"><a:t>1 &amp; 2</a:t><a:t>3</a:t><a:e k="v"/></a:root>';
		const tree = inkObjectFromXml(xml);
		expect(inkObjectToXml(tree)).toBe(
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:root xmlns:a="urn:a" note="&quot;x&lt;"><a:t>1 &amp; 2</a:t><a:t>3</a:t><a:e k="v"/></a:root>',
		);
	});
});

describe('inkSvg', () => {
	it('draws checked paths and colours only', () => {
		const svg = inkSvg(
			[
				{ path: 'M 1 1 L 5 5', color: '#ff0000', width: 2, opacity: 0.4 },
				{ path: 'M 1 1"/><script>', color: '#000000', width: 1, opacity: 1 },
				{ path: 'M 0 0 L 1 1', color: 'red" onload="x', width: 1, opacity: 1 },
			],
			10,
			10,
			20,
			20,
		);
		expect(svg).toContain('stroke="#ff0000"');
		expect(svg).toContain('stroke-opacity="0.4"');
		expect(svg).toContain('viewBox="0 0 20 20"');
		expect(svg).toContain('vector-effect="non-scaling-stroke"');
		expect(svg).not.toContain('script');
		expect(svg).not.toContain('onload');
	});
});
