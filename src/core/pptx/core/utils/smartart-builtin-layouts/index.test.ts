import { readFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../PptxHandler';
import type { PptxElement, SmartArtPptxElement } from '../../types';
import { computeSmartArtElementsWithoutCache, decomposeSmartArt } from '../index';
import {
	applyBuiltinSmartArtLayout,
	findBuiltinSmartArtLayout,
	listBuiltinSmartArtLayouts,
	loadBuiltinSmartArtLayoutXml,
} from './index';

const GALLERY = path.resolve(
	path.dirname(fileURLToPath(import.meta.url)),
	'../../../__tests__/fixtures/smartart-gallery',
);

async function load(file: string): Promise<SmartArtPptxElement> {
	const buf = readFileSync(path.join(GALLERY, file));
	const { slides } = await new PptxHandler().load(
		buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer,
	);
	const element = slides
		.flatMap((slide) => slide.elements)
		.find((el): el is SmartArtPptxElement => el.type === 'smartArt');
	if (!element?.smartArtData) throw new Error(`no SmartArt in ${file}`);
	return element;
}

const textKeyed = (elements: PptxElement[]) => {
	const map = new Map<string, Extract<PptxElement, { type: 'shape' }>>();
	for (const el of elements) {
		if (el.type !== 'shape') continue;
		const text = (el.text ?? '').trim();
		if (
			text &&
			(!map.has(text) || el.width * el.height > map.get(text)!.width * map.get(text)!.height)
		) {
			map.set(text, el);
		}
	}
	return map;
};

describe('built-in SmartArt layout catalogue', () => {
	it('lists the 176 PowerPoint layouts with titles and categories', () => {
		const layouts = listBuiltinSmartArtLayouts();
		expect(layouts).toHaveLength(176);
		expect(new Set(layouts.map((layout) => layout.id)).size).toBe(176);
		expect(layouts.every((layout) => layout.title && layout.category)).toBe(true);
		expect(layouts.map((layout) => layout.title)).toEqual(
			expect.arrayContaining(['Basic Timeline', 'Circle Relationship', 'Basic Bending Process']),
		);
	});

	it('inflates a layout definition on demand and rejects an unknown id', async () => {
		const entry = listBuiltinSmartArtLayouts().find((layout) => layout.title === 'Basic Timeline')!;
		expect(await loadBuiltinSmartArtLayoutXml(entry.id)).toMatch(/<[\w:]*layoutDef\b/u);
		await expect(loadBuiltinSmartArtLayoutXml('urn:nope')).rejects.toThrow(/Unknown built-in/u);
		expect(findBuiltinSmartArtLayout('urn:nope')).toBeUndefined();
	});
});

describe('swapping to a built-in layout runs the real engine', () => {
	// A swap keeps the nodes and drops the old definition (`switchSmartArtLayout`). Starting from the
	// target fixture's own nodes with its definition and drawing cleared, applying the built-in
	// layout must reproduce PowerPoint's cached drawing of that same data.
	const targets = [
		['Circle Relationship', 'circle-relationship--hier5.pptx'],
		['Basic Timeline', 'basic-timeline--hier5.pptx'],
		['Basic Bending Process', 'basic-bending-process--hier5.pptx'],
	] as const;

	for (const [title, file] of targets) {
		it(`${title}: every text shape lands within 1% of PowerPoint's own drawing`, async () => {
			const target = await load(file);
			const source = {
				...target.smartArtData!,
				layoutDefinition: undefined,
				presLayoutVars: undefined,
				drawingShapes: [],
			};
			const entry = listBuiltinSmartArtLayouts().find((layout) => layout.title === title)!;
			const swapped = await applyBuiltinSmartArtLayout(source, entry.id);
			expect(swapped.layoutDefinition?.uniqueId).toBe(entry.id);

			const bounds = { x: target.x, y: target.y, width: target.width, height: target.height };
			const cached = textKeyed(decomposeSmartArt(target.smartArtData!, bounds) ?? []);
			const engine = textKeyed(computeSmartArtElementsWithoutCache(swapped, bounds) ?? []);
			expect(cached.size).toBeGreaterThan(0);
			expect(engine.size).toBe(cached.size);
			for (const [text, expected] of cached) {
				const actual = engine.get(text);
				expect(actual, text).toBeDefined();
				expect(actual!.shapeType, text).toBe(expected.shapeType);
				for (const [a, b, whole] of [
					[actual!.x, expected.x, bounds.width],
					[actual!.y, expected.y, bounds.height],
					[actual!.width, expected.width, bounds.width],
					[actual!.height, expected.height, bounds.height],
				] as const) {
					expect(Math.abs(a - b) / whole, text).toBeLessThan(0.01);
				}
			}
		});
	}
});
