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

	it('marks a part read only in part, and never rewrites it', () => {
		const head =
			'<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML"><inkml:traceFormat><inkml:channel name="X" type="integer"/><inkml:channel name="Y" type="integer"/></inkml:traceFormat>';
		const big = Array.from({ length: 5000 }, (_, i) => `${i} ${i}`).join(',');
		const xml = `${head}<inkml:trace>0 0,10 10</inkml:trace><inkml:trace>${big}</inkml:trace></inkml:ink>`;
		const part = readInkPart(xml, { width: 100, height: 100 });
		expect(part.truncated).toBe(true);
		expect(part.strokes).toHaveLength(1);
		// An edit of the strokes read leaves the part alone: the long trace would be lost.
		expect(updatedInkPart(xml, [], { width: 100, height: 100 })).toBeUndefined();
		expect(readInkPart('<ink:ink').truncated).toBe(true);
		expect(readInkPart('x'.repeat(9 * 1024 * 1024)).truncated).toBe(true);
		expect(readInkPart(`${head}<inkml:trace>0 0,10 10</inkml:trace></inkml:ink>`).truncated).toBe(
			false,
		);
	});

	it('reads pressure by the declared range and InkML ?, * and # values', () => {
		const xml =
			'<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML"><inkml:traceFormat><inkml:channel name="X" type="integer"/><inkml:channel name="Y" type="integer"/><inkml:channel name="F" type="integer" max="1023"/></inkml:traceFormat>' +
			"<inkml:trace>#A 0 1023,'1 ? *,* * 0</inkml:trace></inkml:ink>";
		const [stroke] = readInkPart(xml).strokes;
		expect(stroke?.pressures).toEqual([1, 1, 0]);
		// #A is 10; the second point moves by 1 and keeps y; * repeats the last difference.
		expect(stroke?.path).toBe('M 10 0 L 11 0 L 12 0');
	});

	it('keeps brush properties it does not model and writes Office’s highlighter', () => {
		const xml =
			'<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML"><inkml:definitions><inkml:brush xml:id="b"><inkml:brushProperty name="width" value="0.1" units="cm"/><inkml:brushProperty name="height" value="0.3" units="cm"/><inkml:brushProperty name="color" value="#FF0000"/><inkml:brushProperty name="ignorePressure" value="1"/><inkml:brushProperty name="fitToCurve" value="1"/></inkml:brush></inkml:definitions>' +
			'<inkml:trace brushRef="#b">0 0,10 10</inkml:trace><inkml:trace brushRef="#b">5 5,20 0</inkml:trace></inkml:ink>';
		const box = { width: 20, height: 10 };
		const { strokes } = readInkPart(xml, box);
		const out = updatedInkPart(xml, strokes.slice(0, 1), box)!;
		for (const property of [
			'name="ignorePressure" value="1"',
			'name="fitToCurve" value="1"',
			'name="height" value="0.3" units="cm"',
		])
			expect(out).toContain(property);
		const highlighter = inkStrokeFromSamples(
			[
				{ x: 0, y: 0 },
				{ x: 9, y: 3 },
			],
			{ kind: 'highlighter', color: '#FFFF00', width: 22 },
		)!;
		const written = writeInkPart([highlighter]);
		expect(written).toContain('name="tip" value="rectangle"');
		expect(written).toContain('name="rasterOp" value="maskPen"');
		expect(readInkPart(written).strokes[0]?.tool).toBe('highlighter');
	});

	it('stretches its own strokes onto a box resized after saving', () => {
		const stroke = inkStrokeFromSamples(
			[
				{ x: 0, y: 0 },
				{ x: 50, y: 20 },
			],
			pen,
		)!;
		const xml = writeInkPart([stroke], undefined, { width: 50, height: 20 });
		expect(readInkPart(xml, { width: 50, height: 20 }).strokes[0]?.path).toBe(stroke.path);
		expect(readInkPart(xml, { width: 100, height: 40 }).strokes[0]?.path).toBe('M 0 0 L 100 40');
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

describe('inkStrokesEqual', () => {
	const base = { path: 'M 0 0 L 4 4', color: '#000000', width: 2, opacity: 1 };
	it('tells strokes apart by pressure, tilt, pen and brush properties', () => {
		expect(inkStrokesEqual([{ ...base, pressures: [0.5, 0.5] }], [base])).toBe(true);
		expect(
			inkStrokesEqual([{ ...base, pressures: [0.2, 0.9] }], [{ ...base, pressures: [0.9, 0.2] }]),
		).toBe(false);
		expect(
			inkStrokesEqual([{ ...base, tiltAngles: [1, 2] }], [{ ...base, tiltAngles: [1, 3] }]),
		).toBe(false);
		expect(inkStrokesEqual([{ ...base, tool: 'highlighter' }], [base])).toBe(false);
		expect(
			inkStrokesEqual([{ ...base, brushExtras: [{ name: 'fitToCurve', value: '1' }] }], [base]),
		).toBe(false);
		expect(
			inkStrokesEqual(
				[{ ...base, tool: 'highlighter' }],
				[
					{
						...base,
						tool: 'highlighter',
						brushExtras: [
							{ name: 'tip', value: 'rectangle' },
							{ name: 'rasterOp', value: 'maskPen' },
						],
					},
				],
			),
		).toBe(true);
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
