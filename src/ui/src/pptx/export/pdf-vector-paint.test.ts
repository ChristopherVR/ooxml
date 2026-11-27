// @vitest-environment happy-dom
import { inflateSync } from 'node:zlib';
import { describe, expect, it } from 'vitest';
import { PdfFontSet } from 'ooxml-core/pdf/write';
import { buildVectorPdf } from 'ooxml-core/pdf/write';
import { syntheticTrueType } from '../../../../core/text/truetype/test-support/synthetic-truetype';
import type { DisplayItem, TextItem } from './pdf-vector-display-list';
import { IDENTITY } from './pdf-vector-display-list';
import { paintSlide, resolvePdfTextFont } from './pdf-vector-paint';
import { selectPdfSlides, slidePdfPageSize } from './pdf-vector-select';
import { needsShaping } from './pdf-vector-text';

describe('text that needs shaping', () => {
	it('flags right-to-left, joining and combining text, not plain Latin or CJK', () => {
		expect(needsShaping('Hello, world')).toBe(false);
		expect(needsShaping('中文')).toBe(false);
		expect(needsShaping('שלום')).toBe(true);
		expect(needsShaping('مرحبا')).toBe(true);
		expect(needsShaping('é')).toBe(true);
	});
});

const latin1 = (bytes: Uint8Array) => Buffer.from(bytes).toString('latin1');

/** Every stream of the file, inflated when it is Flate-compressed. */
function streams(bytes: Uint8Array): string[] {
	const text = latin1(bytes);
	const out: string[] = [];
	for (const match of text.matchAll(/<<([^]*?)>>\nstream\n/g)) {
		const length = Number(/\/Length (\d+)/.exec(match[1]!)?.[1]);
		const start = match.index! + match[0].length;
		const data = bytes.subarray(start, start + length);
		out.push(
			match[1]!.includes('/FlateDecode') ? latin1(inflateSync(Buffer.from(data))) : latin1(data),
		);
	}
	return out;
}

const text = (overrides: Partial<TextItem> = {}): TextItem => ({
	kind: 'text',
	matrix: IDENTITY,
	clips: [],
	alpha: 1,
	text: 'Hello world',
	x: 100,
	baseline: 200,
	width: 300,
	size: 40,
	family: '"Calibri", sans-serif',
	bold: false,
	italic: false,
	color: { r: 0, g: 0, b: 0, a: 1 },
	...overrides,
});

const size = { canvasWidth: 1280, canvasHeight: 720, pageWidth: 960, pageHeight: 540 };

describe('vector PDF slide selection', () => {
	const slides = [{}, { hidden: true }, {}, {}];
	it('leaves hidden slides out unless asked for', () => {
		expect(selectPdfSlides(slides)).toEqual([0, 2, 3]);
		expect(selectPdfSlides(slides, { includeHiddenSlides: true })).toEqual([0, 1, 2, 3]);
	});
	it('counts a range over every slide, hidden ones included', () => {
		expect(selectPdfSlides(slides, { from: 2, to: 3 })).toEqual([2]);
		expect(selectPdfSlides(slides, { from: 2, to: 3, includeHiddenSlides: true })).toEqual([1, 2]);
		expect(selectPdfSlides(slides, { from: 0, to: 99 })).toEqual([0, 2, 3]);
		expect(selectPdfSlides(slides, { slideIndexes: [3, 3, 0, 9, -1] })).toEqual([0, 3]);
	});
	it('sizes pages from the slide size, not a paper size', () => {
		expect(
			slidePdfPageSize({ width: 1280, height: 720 }, { widthEmu: 12192000, heightEmu: 6858000 }),
		).toEqual({ width: 960, height: 540 });
		expect(slidePdfPageSize({ width: 960, height: 720 })).toEqual({ width: 720, height: 540 });
	});
});

describe('vector PDF painting', () => {
	it('writes text as real, scaled text in a standard font when no font file is supplied', () => {
		const { page, diagnostics } = paintSlide([text()], size);
		expect(page.width).toBe(960);
		expect(page.height).toBe(540);
		expect(page.content).toMatch(
			/BT \/F0 40 Tf 1 0 0 -1 100 200 Tm .*Tz \(Hello world\) Tj ET 100 Tz/,
		);
		expect([...page.fonts.values()]).toEqual(['Helvetica']);
		expect(diagnostics).toEqual([
			{ code: 'fontStandard', params: { family: 'Calibri', substitute: 'Helvetica' } },
		]);
		const pdf = latin1(buildVectorPdf([page]));
		expect(pdf).toContain('/MediaBox [0 0 960 540]');
		expect(streams(buildVectorPdf([page])).join('\n')).toContain('(Hello world) Tj');
	});

	it('embeds and subsets a supplied font of the run family, with a ToUnicode map', () => {
		const fonts = new PdfFontSet([{ family: 'Calibri', data: syntheticTrueType() }]);
		const item = text({ text: 'AB BA' });
		expect(resolvePdfTextFont(item.text, item.family, false, false, fonts)?.kind).toBe('embedded');
		const { page, diagnostics } = paintSlide([item], { ...size, fonts });
		expect(diagnostics).toEqual([]);
		expect(page.content).toMatch(/\/E0 40 Tf .*TJ ET/);
		const bytes = buildVectorPdf([page]);
		const pdf = latin1(bytes);
		expect(pdf).toContain('/Subtype /Type0');
		expect(pdf).toContain('/Subtype /CIDFontType2');
		expect(pdf).toMatch(/\/BaseFont \/[A-Z]{6}\+Calibri/);
		expect(pdf).toContain('/FontFile2');
		expect(streams(bytes).join('\n')).toContain('beginbfchar');
	});

	it('keeps text over a picture searchable but invisible', () => {
		const { page } = paintSlide([text({ invisible: true })], size);
		expect(page.content).toContain('3 Tr');
		expect(page.content).toContain('0 Tr');
	});

	it('paints paths, gradient shadings, pictures and clips in order', () => {
		const items: DisplayItem[] = [
			{
				kind: 'path',
				matrix: IDENTITY,
				clips: [],
				alpha: 1,
				d: 'M 0 0 L 10 0 L 10 10 Z',
				fill: { r: 1, g: 0, b: 0, a: 1 },
			},
			{
				kind: 'path',
				matrix: [1, 0, 0, 1, 5, 5],
				clips: [{ matrix: IDENTITY, d: 'M 0 0 L 50 0 L 50 50 Z' }],
				alpha: 0.5,
				d: 'M 0 0 L 20 0 L 20 20 Z',
				shading: {
					shading: {
						kind: 'axial',
						coords: [0, 0, 20, 0],
						stops: [
							{ offset: 0, r: 0, g: 0, b: 1 },
							{ offset: 1, r: 1, g: 1, b: 1 },
						],
					},
					matrices: [IDENTITY],
					opacity: 1,
				},
				stroke: { color: { r: 0, g: 0, b: 0, a: 1 }, width: 2, dash: [6, 6], cap: 0, join: 0 },
			},
			{
				kind: 'image',
				matrix: IDENTITY,
				clips: [],
				alpha: 1,
				x: 10,
				y: 20,
				width: 30,
				height: 40,
				image: { width: 1, height: 1, data: new Uint8Array([0, 0, 0]) },
			},
		];
		const { page } = paintSlide(items, size);
		const content = page.content;
		expect(content.indexOf('1 0 0 rg')).toBeLessThan(content.indexOf('/Sh0 sh'));
		expect(content.indexOf('/Sh0 sh')).toBeLessThan(content.indexOf('/Im0 Do'));
		expect(content).toContain('[6 6] 0 d');
		expect(content).toContain('W n');
		expect(content).toContain('30 0 0 -40 10 60 cm');
		expect(page.alphas.size).toBeGreaterThan(0);
		const pdf = latin1(buildVectorPdf([page]));
		expect(pdf).toContain('/ShadingType 2');
		expect(pdf).toContain('/Subtype /Image');
	});

	it('draws inline SVG through the shared converter', () => {
		const svg = new DOMParser().parseFromString(
			'<svg xmlns="http://www.w3.org/2000/svg"><circle cx="5" cy="5" r="4" fill="#00ff00"/></svg>',
			'image/svg+xml',
		).documentElement;
		const { page } = paintSlide(
			[{ kind: 'svg', matrix: IDENTITY, clips: [], alpha: 1, svg, images: new Map() }],
			size,
		);
		expect(page.content).toContain('0 1 0 rg');
		expect(page.content).toMatch(/ c\n/);
	});
});
