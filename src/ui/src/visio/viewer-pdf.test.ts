import { describe, expect, it } from 'vitest';
import { demoDocument } from 'ooxml-core/visio/ui';
import { exportPageSvg } from './export-svg';
import { buildVectorPdf, PdfVectorPage } from '../export/pdf-vector';
import { convertSvgPage } from './pdf-svg';
import { PdfFontSet } from '../export/pdf-embedded-font';
import { existsSync, readFileSync } from 'node:fs';
import { exportDocumentPdfDetailed, pdfPageIndexes } from './viewer-pdf';

const latin1 = (bytes: Uint8Array) => String.fromCharCode(...bytes);
const scene = () => {
	const model = structuredClone(demoDocument);
	// U+2212 has no WinAnsi code and would (correctly) force the picture fallback.
	for (const page of model.pages)
		for (const shape of page.shapes) shape.text.plainText = shape.text.plainText.replace('−', '-');
	return model;
};
const parse = (svg: string) =>
	new DOMParser().parseFromString(svg, 'image/svg+xml').documentElement;
const convert = (model = scene(), index = 0) => {
	const exported = exportPageSvg(model, index);
	return convertSvgPage(parse(exported.svg), exported.width, exported.height, new Map());
};
/** Text drawn by `Tj` operators, in order. */
const shown = (content: string) =>
	[...content.matchAll(/\((.*)\) Tj/g)].map((match) => match[1]!.replace(/\\(.)/g, '$1'));

describe('vector page conversion', () => {
	it('draws shapes as vector paths and text as real text', () => {
		const result = convert();
		expect(result.fallback).toEqual([]);
		const content = result.page!.content;
		expect(content).toMatch(/ m\n/);
		expect(content).toMatch(/ l\n/);
		expect(content).toMatch(/^[\d.]+ [\d.]+ [\d.]+ rg$/m);
		expect(content).toMatch(/^f$|^B$|^S$/m);
		const words = shown(content);
		expect(words).toContain('Start with an idea');
		expect(words).toContain('Review and release');
		expect(words.join(' ')).toContain('Select shapes');
		expect(result.page!.fonts.get('F0')).toBe('Helvetica');
	});

	it('draws stroke weight, dash and arrowheads', () => {
		const model = scene();
		const connector = model.pages[0]!.shapes[1]!;
		connector.style.endArrow = 2;
		connector.style.linePattern = 2;
		connector.style.lineDash = [7, 5];
		connector.style.lineWidth = 0.03;
		const content = convert(model).page!.content;
		expect(content).toMatch(/0\.03 w \d J \d j \[0\.21 0\.15\] 0 d/);
		// The arrowhead glyph is a filled marker path placed with its own transform.
		expect(content.match(/ cm\n/g)!.length).toBeGreaterThan(10);
		expect(content).toMatch(/ l\nf\n/);
	});

	it('lists effects it cannot draw instead of dropping them silently', () => {
		const model = scene();
		const shape = model.pages[0]!.shapes[0]!;
		shape.style.shadow = {
			color: '#000000',
			offsetX: 0.05,
			offsetY: -0.05,
			blur: 0.02,
			opacity: 0.5,
		};
		shape.style.fillGradient = {
			type: 'linear',
			start: [0, 0],
			end: [1, 0],
			stops: [
				{ offset: 0, color: '#ff0000', opacity: 1 },
				{ offset: 1, color: '#0000ff', opacity: 1 },
			],
		} as never;
		const result = convert(model);
		expect(result.fallback).toEqual([]);
		// The gradient is a real shading and the shadow an offset copy; only the blur is lost.
		expect(result.degraded).toEqual(['Shadow blur is not applied: shadows are drawn hard-edged.']);
		expect(result.page!.shadings.size).toBeGreaterThan(0);
		expect(result.page!.content).toMatch(/\/Sh0 sh/);
	});

	it('writes gradients as axial and radial shadings with a stitching function', () => {
		const svg = parse(
			`<svg xmlns="http://www.w3.org/2000/svg"><defs>
				<linearGradient id="l" x1="0" y1="0" x2="1" y2="0"><stop offset="0" stop-color="#ff0000"/><stop offset="0.5" stop-color="#00ff00"/><stop offset="1" stop-color="#0000ff"/></linearGradient>
				<radialGradient id="r" cx="0.5" cy="0.5" r="0.5"><stop offset="0" stop-color="#ffffff"/><stop offset="1" stop-color="#000000" stop-opacity="0.5"/></radialGradient>
			</defs>
			<path d="M 0 0 L 2 0 L 2 1 L 0 1 Z" fill="url(#l)"/>
			<path d="M 0 1 L 2 1 L 2 2 L 0 2 Z" fill="url(#r)"/></svg>`,
		);
		const result = convertSvgPage(svg, 2, 2, new Map());
		expect(result.fallback).toEqual([]);
		expect(result.degraded).toEqual([
			'Gradient transparency is averaged: a PDF shading has one opacity.',
		]);
		const text = latin1(buildVectorPdf([result.page!]));
		expect(text).toContain('/ShadingType 2');
		expect(text).toContain('/ShadingType 3');
		expect(text).toContain('/FunctionType 3');
		expect(text).toMatch(/\/Shading << \/Sh0 << /);
		// Bounding-box units map to the path before the shading is painted inside its clip.
		expect(result.page!.content).toMatch(/W n\n2 0 0 1 0 0 cm\n\/Sh0 sh/);
	});

	it('approximates glow with a halo and reports it', () => {
		const model = scene();
		model.pages[0]!.shapes[0]!.style.glow = { color: '#ffcc00', size: 0.1, opacity: 0.6 } as never;
		const result = convert(model);
		expect(result.fallback).toEqual([]);
		expect(result.degraded.join(' ')).toMatch(/Glow is approximated/);
		expect(result.page!.content).toMatch(/^0\.115 w 1 J 1 j/m);
	});

	it('falls back for content the vector path cannot represent', () => {
		const svg = parse(
			'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 2 2"><g><foreignObject/></g></svg>',
		);
		const result = convertSvgPage(svg, 2, 2, new Map());
		expect(result.page).toBeUndefined();
		expect(result.fallback[0]).toMatch(/foreignObject/);
		const pattern = parse(
			'<svg xmlns="http://www.w3.org/2000/svg"><defs><pattern id="p"/></defs><path d="M 0 0 L 1 1" fill="url(#p)"/></svg>',
		);
		expect(convertSvgPage(pattern, 1, 1, new Map()).fallback[0]).toMatch(/pattern fill/);
		// Text outside the standard-font character set cannot be written as real text.
		const unicode = parse(
			'<svg xmlns="http://www.w3.org/2000/svg"><text><tspan x="0" y="1">Zoom − out</tspan></text></svg>',
		);
		expect(convertSvgPage(unicode, 1, 1, new Map()).fallback[0]).toMatch(/WinAnsi/);
		expect(convertSvgPage(unicode, 1, 1, new Map(), new PdfFontSet([])).fallback[0]).toMatch(
			/WinAnsi/,
		);
	});

	it('embeds decoded pictures', () => {
		const href = 'data:image/png;base64,AAAA';
		const svg = parse(
			`<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"><image href="${href}" width="2" height="1"/></svg>`,
		);
		const resource = { width: 1, height: 1, data: new Uint8Array([1, 2, 3]) };
		const result = convertSvgPage(svg, 2, 1, new Map([[href, resource]]));
		expect(result.page!.content).toContain('/Im0 Do');
		const text = latin1(buildVectorPdf([result.page!]));
		expect(text).toContain('/Subtype /Image');
		expect(text).toContain('/Im0');
	});
});

describe('exportDocumentPdfDetailed', () => {
	it('writes one correctly sized page per foreground page with extractable text', async () => {
		const model = scene();
		const { blob, pages } = await exportDocumentPdfDetailed(document, model);
		const text = latin1(new Uint8Array(await blob.arrayBuffer()));
		expect(text.startsWith('%PDF-1.4')).toBe(true);
		expect(text).toContain('/Count 2');
		expect(text).toContain('/MediaBox [0 0 612 504]');
		expect(text).toContain('/MediaBox [0 0 612 396]');
		expect(pages.map((page) => page.mode)).toEqual(['vector', 'vector']);
		expect(shown(text)).toContain('Framework-neutral rendering');
		expect(text).toMatch(/\/BaseFont \/Helvetica /);
	});

	it('honours the current-page and range options', async () => {
		const model = scene();
		expect(pdfPageIndexes(model, { scope: 'current', pageIndex: 1 })).toEqual([1]);
		expect(pdfPageIndexes(model, { from: 2, to: 2 })).toEqual([1]);
		expect(pdfPageIndexes(model, { scope: 'current', pageIndex: 9 })).toEqual([]);
		const { blob, pages } = await exportDocumentPdfDetailed(document, model, {
			scope: 'current',
			pageIndex: 1,
		});
		expect(pages.map((page) => page.name)).toEqual(['Architecture']);
		expect(latin1(new Uint8Array(await blob.arrayBuffer()))).toContain('/Count 1');
	});
});

describe('PdfVectorPage', () => {
	it('rejects invalid page sizes and shares fonts and alpha states', () => {
		expect(() => new PdfVectorPage(0, 10)).toThrow();
		const page = new PdfVectorPage(100, 100);
		page.alpha(0.5);
		page.alpha(0.5);
		page.text('Helvetica', 12, 0, 0, 'a');
		page.text('Helvetica', 12, 0, 0, 'b');
		expect(page.alphas.size).toBe(1);
		expect(page.fonts.size).toBe(1);
		const text = latin1(buildVectorPdf([page]));
		const xref = Number(/startxref\n(\d+)/.exec(text)![1]);
		expect(text.slice(xref, xref + 4)).toBe('xref');
	});
});

describe('embedded fonts', () => {
	const arial = 'C:/Windows/Fonts/arial.ttf';
	const unicode = parse(
		'<svg xmlns="http://www.w3.org/2000/svg"><text><tspan x="0" y="1" font-family="Arial" font-size="0.2">Zoom − out 中</tspan></text></svg>',
	);
	it('refuses fonts that are not TrueType', async () => {
		const set = new PdfFontSet([{ family: 'Bad', data: new Uint8Array(64) }]);
		expect(set.size).toBe(0);
		const { refusedFonts } = await exportDocumentPdfDetailed(document, scene(), {
			fonts: [{ family: 'Bad', data: new Uint8Array(64) }],
		});
		expect(refusedFonts).toEqual(['Bad']);
	});

	it.skipIf(!existsSync(arial))(
		'keeps text outside WinAnsi as selectable text in a subset font',
		() => {
			const fonts = new PdfFontSet([
				{ family: 'Arial', data: new Uint8Array(readFileSync(arial)) },
			]);
			// Arial has the minus sign but no CJK ideograph: that run still needs a picture.
			expect(convertSvgPage(unicode, 1, 1, new Map(), fonts).fallback[0]).toMatch(/WinAnsi/);
			const minus = parse(
				'<svg xmlns="http://www.w3.org/2000/svg"><text><tspan x="0" y="1" font-family="Arial" font-size="0.2">Zoom − out</tspan></text></svg>',
			);
			const result = convertSvgPage(minus, 1, 1, new Map(), fonts);
			expect(result.fallback).toEqual([]);
			expect(result.page!.content).toMatch(/\/E0 [\d.]+ Tf .* TJ ET/);
			const text = latin1(buildVectorPdf([result.page!]));
			expect(text).toContain('/Subtype /CIDFontType2');
			expect(text).toContain('/Encoding /Identity-H');
			expect(text).toContain('/FontFile2');
			// ToUnicode maps the minus sign back to U+2212.
			expect(text).toContain('<2212>');
			expect(text.length).toBeLessThan(60_000);
		},
	);
});
