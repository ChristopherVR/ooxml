import { describe, it, expect } from 'vitest';
import { assertViewableDocument } from './scene-validation.js';
import { demoDocument } from './demo-document.js';
import { renderText } from './render-text.js';
import { rasterFixture } from '../tests/raster-fixtures.mjs';

describe('defensive scene limits', () => {
	it.each(
		[[], [1], [0, 0], [1, NaN], [1, 28], [1, -1], [1, 1, 1], [1, 1, 1, 1, 1, 1, 1, 1]].map(
			(lineDash) => ({ lineDash }),
		),
	)('rejects invalid normalized line-dash arrays %j', ({ lineDash }) => {
		const model = structuredClone(demoDocument);
		model.pages[0]!.shapes[0]!.style.lineDash = lineDash;
		expect(() => assertViewableDocument(model)).toThrow(/line dash/);
	});
	it('rejects nonfinite and excessive dimensions before rendering', () => {
		const model = structuredClone(demoDocument);
		model.pages[0]!.width = Infinity;
		expect(() => assertViewableDocument(model)).toThrow('page width');
		model.pages[0]!.width = 1e9;
		expect(() => assertViewableDocument(model)).toThrow('page width');
	});
	it('rejects shape cycles without recursion', () => {
		const model = structuredClone(demoDocument);
		const shape = model.pages[0]!.shapes[0]!;
		shape.children = [shape];
		expect(() => assertViewableDocument(model)).toThrow('cycle');
	});
	it('rejects unnormalized line caps supplied through a host model', () => {
		const model = structuredClone(demoDocument);
		Object.assign(model.pages[0]!.shapes[0]!.style, { lineCap: 'url(https://invalid.test/cap)' });
		expect(() => assertViewableDocument(model)).toThrow('normalized line cap');
	});
	it('rejects excess path/text sizes and invalid transforms', () => {
		const model = structuredClone(demoDocument);
		model.pages[0]!.shapes[0]!.transform = [1, 0, 0, NaN, 0, 0];
		expect(() => assertViewableDocument(model)).toThrow('transform');
	});
	it('computes text size with reduction rather than variadic argument expansion', () => {
		const text = structuredClone(demoDocument.pages[0]!.shapes[0]!.text);
		text.runs = Array.from({ length: 1000 }, () => ({
			text: '',
			fontFamily: 'Arial',
			fontSize: 0.2,
			color: '#000',
			bold: false,
			italic: false,
			underline: false,
		}));
		expect(() => renderText(text, new Set())).not.toThrow();
	});
});

describe('aggregate raster and metadata limits', () => {
	it('validates host-supplied raster bytes instead of trusting MIME and dimensions', () => {
		const model = structuredClone(demoDocument),
			shape = model.pages[0]!.shapes[0]!;
		shape.image = rasterFixture();
		assertViewableDocument(model);
		shape.image.mimeType = 'image/jpeg';
		expect(() => assertViewableDocument(model)).toThrow('do not match');
		shape.image.mimeType = 'image/png';
		shape.image.pixelWidth = 2;
		expect(() => assertViewableDocument(model)).toThrow('do not match');
		shape.image.pixelWidth = 1;
		shape.image.bytes = Uint8Array.from(
			new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg" onload="alert(1)"/>'),
		);
		expect(() => assertViewableDocument(model)).toThrow('not a supported');
	});
	it('revalidates mutated byte buffers on every call and catches conflicting shared declarations', () => {
		const model = structuredClone(demoDocument),
			shape = model.pages[0]!.shapes[0]!;
		shape.image = rasterFixture();
		assertViewableDocument(model);
		shape.image.bytes[0] = 0;
		expect(() => assertViewableDocument(model)).toThrow('not a supported');
		shape.image = rasterFixture();
		model.pages[0]!.shapes.push({
			...shape,
			id: 'different-instance',
			image: { ...shape.image, pixelWidth: 2 },
		});
		expect(() => assertViewableDocument(model)).toThrow('inconsistent declared metadata');
	});
	it('bounds decoded pixel cost across repeated image instances', () => {
		const model = structuredClone(demoDocument),
			base = model.pages[0]!.shapes[0]!;
		base.image = {
			mimeType: 'image/png',
			bytes: new Uint8Array([1]),
			pixelWidth: 4000,
			pixelHeight: 4000,
		};
		model.pages[0]!.shapes = Array.from({ length: 17 }, (_, i) => ({ ...base, id: String(i) }));
		expect(() => assertViewableDocument(model)).toThrow('aggregate decoded image');
	});
	it('rejects long font strings and labels before text measurement', () => {
		const model = structuredClone(demoDocument);
		model.pages[0]!.shapes[0]!.text.fontFamily = 'x'.repeat(1025);
		expect(() => assertViewableDocument(model)).toThrow('metadata string');
	});
});

describe('normalized inherited identities', () => {
	it('accepts synthetic inherited shape IDs within the core 1024-character limit', () => {
		const model = structuredClone(demoDocument);
		model.pages[0]!.shapes[0]!.id = 'm'.repeat(512);
		expect(() => assertViewableDocument(model)).not.toThrow();
	});
	it('rejects duplicate page IDs and shape IDs within a page', () => {
		const model = structuredClone(demoDocument);
		model.pages[1]!.id = model.pages[0]!.id;
		expect(() => assertViewableDocument(model)).toThrow('duplicate page');
		model.pages[1]!.id = 'other-page';
		const root = model.pages[0]!.shapes[0]!;
		root.children = [{ ...structuredClone(root), children: [] }];
		expect(() => assertViewableDocument(model)).toThrow('duplicate shape');
	});
	it('allows identical shape IDs on different pages', () => {
		const model = structuredClone(demoDocument);
		model.pages[1]!.shapes = structuredClone(model.pages[0]!.shapes);
		expect(() => assertViewableDocument(model)).not.toThrow();
	});
});
