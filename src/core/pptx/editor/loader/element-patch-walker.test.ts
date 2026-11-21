import type { GroupPptxElement, PptxElement, ZoomPptxElement } from '../../index';
import { describe, expect, it } from 'vitest';

import { applyImagePathPatches, walkAndPatchElements } from './element-patch-walker';
import { collectImagePaths } from './load-content-helpers';
import type { ImagePathElement } from './load-content-helpers';

function makePicture(id: string): PptxElement {
	return { id, type: 'picture', x: 0, y: 0, width: 10, height: 10 } as unknown as PptxElement;
}

describe('walkAndPatchElements', () => {
	it('returns the same array reference when the patcher changes nothing', () => {
		const elements = [makePicture('a'), makePicture('b')];
		const result = walkAndPatchElements(elements, (el) => el);
		expect(result).toBe(elements);
	});

	it('rebuilds only the element the patcher actually changes', () => {
		const a = makePicture('a');
		const b = makePicture('b');
		const elements = [a, b];
		const result = walkAndPatchElements(elements, (el) =>
			el.id === 'a' ? ({ ...el, x: 99 } as PptxElement) : el,
		);
		expect(result).not.toBe(elements);
		expect(result[0]).not.toBe(a);
		expect((result[0] as PptxElement).x).toBe(99);
		expect(result[1]).toBe(b);
	});

	it('recurses into group children and rebuilds the group when a child changes', () => {
		const child = makePicture('child-1');
		const group: GroupPptxElement = {
			id: 'group-1',
			type: 'group',
			x: 0,
			y: 0,
			width: 100,
			height: 100,
			children: [child],
		};
		const elements: PptxElement[] = [group];
		const result = walkAndPatchElements(elements, (el) =>
			el.id === 'child-1' ? ({ ...el, x: 42 } as PptxElement) : el,
		);
		expect(result).not.toBe(elements);
		const newGroup = result[0] as GroupPptxElement;
		expect(newGroup).not.toBe(group);
		expect(newGroup.children[0]).not.toBe(child);
		expect((newGroup.children[0] as PptxElement).x).toBe(42);
	});

	it('leaves an unrelated group untouched when none of its children match', () => {
		const child = makePicture('child-1');
		const group: GroupPptxElement = {
			id: 'group-1',
			type: 'group',
			x: 0,
			y: 0,
			width: 100,
			height: 100,
			children: [child],
		};
		const elements: PptxElement[] = [group];
		const result = walkAndPatchElements(elements, (el) => el);
		expect(result).toBe(elements);
		expect(result[0]).toBe(group);
	});
});

describe('applyImagePathPatches', () => {
	it('returns the same array reference when no ref resolves', () => {
		const elements = [makePicture('pic-1')];
		const refs: ImagePathElement[] = [
			{ element: elements[0]!, field: 'imageData', path: 'ppt/media/image1.png' },
		];
		const result = applyImagePathPatches(elements, new Map(), refs);
		expect(result).toBe(elements);
	});

	it('patches the resolved field onto the matching element only', () => {
		const pic1 = makePicture('pic-1');
		const pic2 = makePicture('pic-2');
		const elements = [pic1, pic2];
		const refs: ImagePathElement[] = [
			{ element: pic1, field: 'imageData', path: 'ppt/media/image1.png' },
		];
		const resolvedMap = new Map([['ppt/media/image1.png', 'blob:resolved-1']]);

		const result = applyImagePathPatches(elements, resolvedMap, refs);

		expect(result).not.toBe(elements);
		expect((result[0] as PptxElement & { imageData?: string }).imageData).toBe('blob:resolved-1');
		expect(result[1]).toBe(pic2);
		// The original element is not mutated in place.
		expect((pic1 as PptxElement & { imageData?: string }).imageData).toBeUndefined();
	});

	it('recurses into group children', () => {
		const child = makePicture('child-1');
		const group: GroupPptxElement = {
			id: 'group-1',
			type: 'group',
			x: 0,
			y: 0,
			width: 100,
			height: 100,
			children: [child],
		};
		const elements: PptxElement[] = [group];
		const refs: ImagePathElement[] = [
			{ element: child, field: 'imageData', path: 'ppt/media/child.png' },
		];
		const resolvedMap = new Map([['ppt/media/child.png', 'blob:resolved-child']]);

		const result = applyImagePathPatches(elements, resolvedMap, refs);

		const newGroup = result[0] as GroupPptxElement;
		expect((newGroup.children[0] as PptxElement & { imageData?: string }).imageData).toBe(
			'blob:resolved-child',
		);
	});
});

it('resolves saved Zoom previews and their Summary Zoom tiles through the shared loader', () => {
	const frame = { x: 0, y: 0, width: 200, height: 100 };
	const slideZoom: ZoomPptxElement = {
		...frame,
		type: 'zoom',
		id: 'slide-zoom',
		zoomType: 'slide',
		targetSlideIndex: 1,
		imagePath: 'ppt/media/slide.png',
	};
	const summaryZoom: ZoomPptxElement = {
		...frame,
		type: 'zoom',
		id: 'summary-zoom',
		zoomType: 'summary',
		targetSlideIndex: 1,
		summaryTargets: [
			{ ...frame, sectionId: 'agenda', targetSlideIndex: 1, imagePath: 'ppt/media/slide.png' },
			{ ...frame, sectionId: 'appendix', targetSlideIndex: 3, imagePath: 'ppt/media/appendix.png' },
		],
	};
	const elements: PptxElement[] = [slideZoom, summaryZoom];
	const { paths, refs } = collectImagePaths([
		{ id: 'first', rId: 'rId1', slideNumber: 1, elements },
	]);
	expect([...paths]).toEqual(['ppt/media/slide.png', 'ppt/media/appendix.png']);
	expect(
		applyImagePathPatches(elements, new Map([['ppt/media/slide.png', 'blob:resolved']]), refs),
	).toMatchObject([
		{ id: 'slide-zoom', imageData: 'blob:resolved' },
		{
			id: 'summary-zoom',
			summaryTargets: [
				{ sectionId: 'agenda', imageData: 'blob:resolved' },
				{ sectionId: 'appendix', imagePath: 'ppt/media/appendix.png' },
			],
		},
	]);
	expect(slideZoom.imageData).toBeUndefined();
	expect(summaryZoom.summaryTargets?.[0]?.imageData).toBeUndefined();
});
