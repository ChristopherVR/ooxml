import { describe, expect, it } from 'vitest';
import { demoDocument } from './demo-document.js';
import { copySnapshotScene } from './snapshot-scene.js';
import { assertViewableDocument } from './scene-validation.js';
import { createPrintSnapshot } from './print-snapshot.js';
import { rasterFixture } from '../tests/raster-fixtures.mjs';

function scene() {
	const model = structuredClone(demoDocument);
	model.pages = [model.pages[0]!];
	const shape = model.pages[0]!.shapes[0]!;
	model.pages[0]!.shapes = [shape];
	shape.image = rasterFixture();
	shape.style.lineDash = [1, 2];
	shape.style.fillGradient = {
		type: 'linear',
		start: [0, 0],
		end: [1, 1],
		stops: [
			{ offset: 0, color: '#000', opacity: 1 },
			{ offset: 1, color: '#fff', opacity: 0.5 },
		],
	};
	shape.text.runs = [
		{
			text: shape.text.plainText,
			fontFamily: 'Arial',
			fontSize: 0.1,
			color: '#000',
			bold: true,
			italic: false,
			underline: false,
		},
	];
	shape.text.paragraphs = [
		{
			start: 0,
			end: shape.text.plainText.length,
			horizontalAlign: 'left',
			indentLeft: 0,
			indentRight: 0,
			indentFirst: 0,
			spaceBefore: 0,
			spaceAfter: 0,
			lineSpacing: { kind: 'multiple', value: 1.2 },
			direction: 'ltr',
			bullet: { text: '•', fontFamily: 'Arial', fontSize: 0.1, offset: 0 },
		},
	];
	return model;
}
describe('bounded private rendering-scene copy', () => {
	it('detaches every nested rendering field and copies shared raster bytes exactly once', () => {
		const source = scene(),
			shape = source.pages[0]!.shapes[0]!;
		source.pages[0]!.shapes.push({
			...structuredClone(shape),
			id: 'shared-image',
			image: shape.image!,
		});
		assertViewableDocument(source);
		const copied = copySnapshotScene(source),
			copy = copied.pages[0]!.shapes[0]!;
		assertViewableDocument(copied);
		expect(copied).toEqual(source);
		for (const [first, second] of [
			[source.pages, copied.pages],
			[source.pages[0], copied.pages[0]],
			[shape, copy],
			[shape.transform, copy.transform],
			[shape.children, copy.children],
			[shape.geometry, copy.geometry],
			[shape.geometry[0], copy.geometry[0]],
			[shape.style, copy.style],
			[shape.style.lineDash, copy.style.lineDash],
			[shape.style.fillGradient, copy.style.fillGradient],
			[shape.style.fillGradient!.start, copy.style.fillGradient!.start],
			[shape.style.fillGradient!.end, copy.style.fillGradient!.end],
			[shape.style.fillGradient!.stops, copy.style.fillGradient!.stops],
			[shape.style.fillGradient!.stops[0], copy.style.fillGradient!.stops[0]],
			[shape.text, copy.text],
			[shape.text.transform, copy.text.transform],
			[shape.text.margins, copy.text.margins],
			[shape.text.runs, copy.text.runs],
			[shape.text.runs[0], copy.text.runs[0]],
			[shape.text.paragraphs, copy.text.paragraphs],
			[shape.text.paragraphs![0], copy.text.paragraphs![0]],
			[shape.text.paragraphs![0]!.lineSpacing, copy.text.paragraphs![0]!.lineSpacing],
			[shape.text.paragraphs![0]!.bullet, copy.text.paragraphs![0]!.bullet],
			[shape.image, copy.image],
			[shape.image!.bytes, copy.image!.bytes],
		])
			expect(second).not.toBe(first);
		expect(copied.pages[0]!.shapes[1]!.image!.bytes).toBe(copy.image!.bytes);
		shape.image!.bytes[0] = 0;
		expect(copy.image!.bytes[0]).toBe(137);
	});
	it('never traverses unrelated, cyclic or uncloneable host properties', () => {
		const source = scene(),
			shape = source.pages[0]!.shapes[0]!;
		for (const target of [
			source,
			source.pages[0]!,
			shape,
			shape.style,
			shape.text,
			shape.text.margins,
			shape.image!,
		]) {
			Object.defineProperty(target, 'unknownHostProperty', {
				enumerable: true,
				get() {
					throw new Error('Unknown getter must not run');
				},
			});
			Object.assign(target, { unknownCycle: source, unknownFunction: () => undefined });
		}
		expect(() => createPrintSnapshot(source, { pageIndices: [0] })).not.toThrow();
		const copy = copySnapshotScene(source);
		expect(Object.hasOwn(copy, 'unknownHostProperty')).toBe(false);
		expect(Object.hasOwn(copy.pages[0]!.shapes[0]!, 'unknownCycle')).toBe(false);
	});
	it('accepts the exact canonical depth 64 and rejects depth 65', () => {
		const source = scene(),
			root = source.pages[0]!.shapes[0]!;
		delete root.image;
		root.geometry = [];
		root.text.plainText = '';
		root.text.runs = [];
		root.text.paragraphs = [];
		let leaf = root;
		for (let depth = 1; depth <= 64; depth++) {
			const child = structuredClone(root);
			child.id = `depth-${depth}`;
			child.children = [];
			leaf.children = [child];
			leaf = child;
		}
		assertViewableDocument(source);
		expect(() => createPrintSnapshot(source, { pageIndices: [0] })).not.toThrow();
		leaf.children = [{ ...structuredClone(leaf), id: 'depth-65', children: [] }];
		expect(() => copySnapshotScene(source)).toThrow('copy limits');
		expect(() => createPrintSnapshot(source, { pageIndices: [0] })).toThrow('depth');
	});
	it('bounds source arrays before allocating copies even without relying on prior validation', () => {
		const source = scene(),
			shape = source.pages[0]!.shapes[0]!;
		source.pages[0]!.shapes = Array(25_001).fill(shape);
		expect(() => copySnapshotScene(source)).toThrow('copy limits');
		source.pages[0]!.shapes = [shape];
		shape.geometry = Array(100_001).fill(shape.geometry[0]);
		expect(() => copySnapshotScene(source)).toThrow('copy limits');
	});
});
