import JSZip from 'jszip';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';

import { createDefaultPptxHandlerRuntime } from '../../core/core';
import {
	masterPartElementSignature,
	templateElementSignature,
} from '../../core/core/runtime/master-part-element-signature';
import { resolveTemplateImages } from '../../core/core/runtime/resolve-template-images';
import { PptxHandler } from '../../core/PptxHandler';
import type {
	GroupPptxElement,
	PicturePptxElement,
	PptxElement,
	PptxSlide,
	ShapePptxElement,
} from '../../core/types';
import {
	buildUnusedLayoutImagesDeck,
	UNUSED_LAYOUT_PATH,
	USED_LAYOUT_PATH,
} from '../fixtures/unused-layout-images';

function arrayBuffer(bytes: Uint8Array): ArrayBuffer {
	return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

function picture(elements: PptxElement[], name: string): PicturePptxElement {
	const element = elements.find((item) => item.name === name);
	expect(element?.type).toBe('picture');
	return element as PicturePptxElement;
}

function syntheticSlide(): PptxSlide {
	return { id: 'slide-new', rId: '', slideNumber: 2, elements: [] };
}

function attachTemplate(slide: PptxSlide, template: PptxElement[]): PptxSlide {
	return {
		...slide,
		elements: [
			...template,
			...slide.elements.filter((element) => !/^(?:layout|master)-/u.test(element.id)),
		],
	};
}

async function withSvg(bytes: Uint8Array): Promise<Uint8Array> {
	const zip = await JSZip.loadAsync(bytes);
	zip.file(
		USED_LAYOUT_PATH,
		(await zip.file(USED_LAYOUT_PATH)!.async('string')).replace(
			'<a:blip r:embed="rIdLogo"/>',
			'<a:blip r:embed="rIdLogo"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">' +
				'<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rIdSvg"/>' +
				'</a:ext></a:extLst></a:blip>',
		),
	);
	const rels = 'ppt/slideLayouts/_rels/slideLayout1.xml.rels';
	zip.file(
		rels,
		(await zip.file(rels)!.async('string')).replace(
			'</Relationships>',
			'<Relationship Id="rIdSvg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/logo.svg"/></Relationships>',
		),
	);
	zip.file(
		'ppt/media/logo.svg',
		'<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><rect width="10" height="10" fill="orange"/></svg>',
	);
	zip.file(
		'[Content_Types].xml',
		(await zip.file('[Content_Types].xml')!.async('string')).replace(
			'</Types>',
			'<Default Extension="svg" ContentType="image/svg+xml"/></Types>',
		),
	);
	return zip.generateAsync({ type: 'uint8array' });
}

describe('images from layouts not used by a loaded slide', () => {
	let bytes: Uint8Array;
	const handlers: PptxHandler[] = [];
	const newHandler = () => {
		const handler = new PptxHandler();
		handlers.push(handler);
		return handler;
	};
	beforeAll(async () => {
		bytes = await buildUnusedLayoutImagesDeck();
	});
	afterEach(() => {
		for (const handler of handlers.splice(0)) handler.dispose();
		vi.restoreAllMocks();
	});

	for (const insert of [true, false]) {
		it(`${insert ? 'inserts cold' : 'applies after a gallery preview'}, then saves and reopens layout artwork`, async () => {
			const runtime = createDefaultPptxHandlerRuntime();
			const getImageData = vi.spyOn(runtime, 'getImageData');
			const handler = new PptxHandler({ runtime });
			handlers.push(handler);
			const data = await handler.load(arrayBuffer(bytes));
			expect(data.slides).toHaveLength(1);
			expect(data.slides[0].layoutPath).toBe(USED_LAYOUT_PATH);
			expect(getImageData.mock.calls.some(([path]) => path === 'ppt/media/unused-logo.png')).toBe(
				false,
			);
			if (!insert) await handler.getLayoutPreview(UNUSED_LAYOUT_PATH);
			const target = insert ? [syntheticSlide()] : data.slides;
			await handler.applyLayoutToSlide(0, UNUSED_LAYOUT_PATH, target);
			getImageData.mockClear();
			const template = await handler.getTemplateElementsForSlide(target[0].id);
			const logo = picture(template, 'Unused Layout Logo');
			expect(logo.imageData).toBeTruthy();
			const zip = await JSZip.loadAsync(bytes);
			expect(new Uint8Array(await (await fetch(logo.imageData!)).arrayBuffer())).toEqual(
				await zip.file('ppt/media/unused-logo.png')!.async('uint8array'),
			);
			const small = picture(template, 'Small Layout Logo');
			expect([small.width, small.height]).toEqual([11, 13]);
			expect(small.imageData).toBe(logo.imageData);
			expect(await handler.getTemplateElementsForSlide(target[0].id)).toBe(template);
			expect(getImageData.mock.calls.filter(([path]) => path === logo.imagePath)).toHaveLength(1);
			const slides = insert
				? [data.slides[0], attachTemplate(target[0], template)]
				: [attachTemplate(target[0], template)];
			const reopened = newHandler();
			const result = await reopened.load(arrayBuffer(await handler.save(slides)));
			expect(result.slides).toHaveLength(insert ? 2 : 1);
			const last = result.slides[result.slides.length - 1];
			expect(last.layoutPath).toBe(UNUSED_LAYOUT_PATH);
			const inherited = await reopened.getTemplateElementsForSlide(last.id);
			expect(picture(inherited, 'Unused Layout Logo').imageData).toBeTruthy();
			expect(result.slides[0].elements.some((element) => element.name === 'Fixture Title')).toBe(
				true,
			);
			if (insert) {
				const used = await reopened.getTemplateElementsForSlide(result.slides[0].id);
				expect(picture(used, 'Used Layout Logo').imageData).not.toBe(
					picture(inherited, 'Unused Layout Logo').imageData,
				);
			}
		});
	}

	for (const svg of [false, true]) {
		it(`preserves shared ${svg ? 'SVG' : 'PNG'} template identities and edit writeback`, async () => {
			const handler = newHandler();
			const data = await handler.load(arrayBuffer(svg ? await withSvg(bytes) : bytes));
			const first = data.slides[0];
			const logo = picture(first.elements, 'Used Layout Logo');
			const rawXml = logo.rawXml;
			const template = await handler.getTemplateElementsForSlide(first.id);
			expect(picture(template, 'Used Layout Logo')).toBe(logo);
			expect(logo.rawXml).toBe(rawXml);
			expect(logo.imageData).toBeTruthy();
			if (svg) expect(logo.svgData).toBeTruthy();
			await handler.getLayoutPreview(USED_LAYOUT_PATH);
			const sibling = [syntheticSlide()];
			await handler.applyLayoutToSlide(0, USED_LAYOUT_PATH, sibling);
			const siblingTemplate = await handler.getTemplateElementsForSlide(sibling[0].id);
			expect(picture(siblingTemplate, 'Used Layout Logo').rawXml).toBe(rawXml);
			const edited = {
				...first,
				isDirty: true,
				elements: first.elements.map((element) =>
					element === logo ? { ...element, x: 900 } : element,
				),
			};
			const saved = await handler.save([edited, attachTemplate(sibling[0], siblingTemplate)]);
			const result = await newHandler().load(arrayBuffer(saved));
			for (const slide of result.slides) {
				expect(picture(slide.elements, 'Used Layout Logo').x).toBe(900);
			}
		});
	}

	it('resolves nested pictures/SVGs once without changing editable fill URLs', async () => {
		const base = { x: 0, y: 0, width: 11, height: 13 };
		const logo: PicturePptxElement = {
			...base,
			id: 'logo',
			type: 'picture',
			imagePath: 'ppt/logo.png',
			svgPath: 'ppt/logo.svg',
		};
		const missing: PicturePptxElement = {
			...base,
			id: 'missing',
			type: 'picture',
			imagePath: 'ppt/missing.png',
		};
		const loaded: PicturePptxElement = {
			...base,
			id: 'loaded',
			type: 'picture',
			imagePath: 'ppt/loaded.png',
			imageData: 'blob:existing',
		};
		const external: PicturePptxElement = {
			...base,
			id: 'external',
			type: 'picture',
			imagePath: 'https://images.example.invalid/logo.png',
		};
		const fill = { fillImageUrl: 'ppt/logo.png' };
		const shape: ShapePptxElement = { ...base, id: 'shape', type: 'shape', shapeStyle: fill };
		const group: GroupPptxElement = {
			...base,
			id: 'group',
			type: 'group',
			children: [logo, shape, missing, loaded, external],
			groupFill: fill,
		};
		const elements = [group];
		const resolve = vi.fn(async (path: string) =>
			path === 'ppt/missing.png' ? undefined : `blob:${path}`,
		);
		const baseline = templateElementSignature(logo);
		const partBaseline = masterPartElementSignature(elements);
		expect(await resolveTemplateImages(elements, resolve)).toBe(elements);
		expect(templateElementSignature(logo)).toBe(baseline);
		expect(masterPartElementSignature(elements)).toBe(partBaseline);
		expect(templateElementSignature({ ...logo, svgPath: 'ppt/changed.svg' })).not.toBe(baseline);
		expect(logo.imageData).toBe('blob:ppt/logo.png');
		expect(logo.svgData).toBe('blob:ppt/logo.svg');
		expect(shape.shapeStyle?.fillImageUrl).toBe('ppt/logo.png');
		expect(group.groupFill).toBe(fill);
		expect(missing.imageData).toBeUndefined();
		expect(loaded.imageData).toBe('blob:existing');
		expect(external.imageData).toBeUndefined();
		expect(resolve.mock.calls.map(([path]) => path)).toEqual([
			'ppt/logo.png',
			'ppt/logo.svg',
			'ppt/missing.png',
		]);
	});

	it('does not fetch or hydrate an externally linked image when external images are blocked', async () => {
		const zip = await JSZip.loadAsync(bytes);
		const relsPath = 'ppt/slideLayouts/_rels/slideLayout2.xml.rels';
		zip.file(
			relsPath,
			(await zip.file(relsPath)!.async('string')).replace(
				'Target="../media/unused-logo.png"',
				'Target="https://images.example.invalid/logo.png" TargetMode="External"',
			),
		);
		const fetchSpy = vi.spyOn(globalThis, 'fetch');
		const handler = newHandler();
		const data = await handler.load(await zip.generateAsync({ type: 'arraybuffer' }));
		await handler.applyLayoutToSlide(0, UNUSED_LAYOUT_PATH, data.slides);
		const template = await handler.getTemplateElementsForSlide(data.slides[0].id);
		expect(picture(template, 'Unused Layout Logo').imageData).toBeUndefined();
		expect(fetchSpy).not.toHaveBeenCalled();
	});
});
