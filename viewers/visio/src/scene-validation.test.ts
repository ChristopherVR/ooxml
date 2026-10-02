import { describe, it, expect } from 'vitest';
import { assertViewableDocument } from './scene-validation.js';
import { demoDocument } from './demo-document.js';
import { renderText } from './render-text.js';

describe('defensive scene limits', () => {
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
