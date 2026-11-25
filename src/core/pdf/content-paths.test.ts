import { describe, expect, it } from 'vitest';
import { openPdf } from './index';
import { buildPdf, simplePdf, stream } from './test-support/build';

const page = async (content: string, resources = '') =>
	(await openPdf(simplePdf(content, { resources }))).page(0);

describe('painted paths', () => {
	it('returns curves, diagonals and polygons as paths in reading coordinates', async () => {
		const result = await page(
			'1 0 0 rg 100 700 m 150 750 200 750 250 700 c f ' +
				'0 0 1 RG 2 w 100 600 m 200 500 l S ' +
				'0 1 0 rg 300 600 m 400 600 l 350 500 l f* ' +
				'0 0 0 rg 72 72 100 20 re f',
		);
		expect(result.paths).toHaveLength(3);
		const [curve, diagonal, triangle] = result.paths!;
		expect(curve).toMatchObject({ fill: '#ff0000', x0: 100, y0: 42, x1: 250, y1: 92 });
		expect(curve!.segments).toEqual([
			{ op: 'M', x: 100, y: 92 },
			{ op: 'C', x1: 150, y1: 42, x2: 200, y2: 42, x: 250, y: 92 },
		]);
		expect(curve!.stroke).toBeUndefined();
		expect(diagonal).toMatchObject({ stroke: '#0000ff', lineWidth: 2 });
		expect(diagonal!.fill).toBeUndefined();
		expect(triangle).toMatchObject({ fill: '#00ff00', evenOdd: true });
		// The rectangle stays a rule, for tables and shading.
		expect(result.rules).toEqual([
			expect.objectContaining({ kind: 'rect', x0: 72, y0: 700, x1: 172, y1: 720, filled: true }),
		]);
		expect(result.notes).toEqual([]);
	});

	it('keeps outline colours, opacity and dashes, and closes rectangles drawn with lines', async () => {
		const result = await page(
			'/GS1 gs 1 1 0 rg 0 0 1 RG [3 2] 0 d 100 100 50 50 re B ' +
				'[] 0 d 300 300 m 400 300 l 400 350 l 300 350 l h S',
			'/ExtGState << /GS1 << /ca 0.5 >> >>',
		);
		expect(result.rules[0]).toMatchObject({
			kind: 'rect',
			color: '#ffff00',
			strokeColor: '#0000ff',
			alpha: 0.5,
			dash: [3, 2],
		});
		// Four edges: the closing one included.
		expect(result.rules.slice(1).map((rule) => rule.kind)).toEqual([
			'line',
			'line',
			'line',
			'line',
		]);
		expect(result.paths).toEqual([]);
	});

	it('clips rectangular clips and reports unsupported shadings, patterns and non-rectangular clips', async () => {
		const result = await page(
			'q 0 0 612 792 re W n 1 0 0 rg 100 100 m 150 150 l 200 100 l f Q ' +
				'q 100 100 m 150 100 l 125 150 l h W n 0 0 1 rg 90 90 m 300 300 l 90 300 l f Q ' +
				'/Pattern cs /P1 scn 100 400 m 150 450 l 200 400 l f ' +
				'/Sh1 sh',
		);
		expect(result.paths).toHaveLength(1);
		expect(result.notes).toEqual([
			'Vector drawings (3 shapes) were left out: a clipping path cuts them, pattern fills are not supported, shadings (gradients) are not supported.',
		]);
	});

	it('gives a form drawing the colour of the content that draws it', async () => {
		const bytes = buildPdf([
			'<< /Type /Catalog /Pages 2 0 R >>',
			'<< /Type /Pages /Kids [3 0 R] /Count 1 /MediaBox [0 0 612 792] >>',
			'<< /Type /Page /Parent 2 0 R /Contents 4 0 R /Resources << /XObject << /Fm 5 0 R >> >> >>',
			stream('', '0 0 1 rg q 1 0 0 1 100 100 cm /Fm Do Q'),
			stream('/Type /XObject /Subtype /Form /BBox [0 0 100 100]', '0 0 m 50 100 l 100 0 l f'),
		]);
		const result = await (await openPdf(bytes)).page(0);
		expect(result.paths).toEqual([
			expect.objectContaining({ fill: '#0000ff', x0: 100, y0: 592, x1: 200, y1: 692 }),
		]);
	});
});
