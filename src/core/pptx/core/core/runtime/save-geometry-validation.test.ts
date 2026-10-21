import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { PptxHandler } from '../../PptxHandler';
import type { PptxElement } from '../../types';
import { PptxElementTransformUpdater } from '../builders/PptxElementTransformUpdater';
import { validateElementGeometry } from './save-geometry-validation';

const TYPES = [
	'text',
	'shape',
	'connector',
	'image',
	'picture',
	'media',
	'table',
	'chart',
	'smartArt',
	'group',
	'contentPart',
	'ink',
	'model3d',
	'zoom',
	'ole',
] as const;
const FIELDS = ['x', 'y', 'width', 'height'] as const;
const INVALID = [undefined, NaN, Infinity, -Infinity];

function element(type: string = 'shape'): PptxElement {
	return {
		id: 'bad-element',
		name: 'Broken frame',
		type,
		x: 10,
		y: 20,
		width: 300,
		height: 200,
		...(type === 'group' ? { children: [] } : {}),
	} as PptxElement;
}

describe('save geometry preflight', () => {
	it.each(TYPES)(
		'rejects every invalid frame field before the %s writer mutates the package',
		async (type) => {
			const { handler, data, createSlide } = await PptxHandler.createBlank();
			const slide = createSlide('Blank').build();
			data.slides.push(slide);
			for (const field of FIELDS) {
				for (const value of INVALID) {
					const invalid = element(type);
					(invalid as unknown as Record<string, unknown>)[field] = value;
					slide.elements = [invalid];
					await expect(handler.save(data.slides)).rejects.toThrow(
						new RegExp(
							`Slide 1 \\(slide1\\).*bad-element.*Broken frame.*type=${type}.*${field}=${String(value)}`,
						),
					);
					expect(slide.id).toBe('slide1');
					expect(invalid.rawXml).toBeUndefined();
				}
			}
		},
	);

	it.each(TYPES)(
		'rejects invalid raw-XML-backed %s geometry without patching XML',
		async (type) => {
			const { handler, data, createSlide } = await PptxHandler.createBlank();
			const slide = createSlide('Blank').build();
			data.slides.push(slide);
			for (const field of FIELDS) {
				for (const value of INVALID) {
					const invalid = element(type);
					invalid.rawXml = {
						'p:spPr': {
							'a:xfrm': {
								'a:off': { '@_x': '95250', '@_y': '190500' },
								'a:ext': { '@_cx': '2857500', '@_cy': '1905000' },
							},
						},
					};
					const baseline = structuredClone(invalid.rawXml);
					(invalid as unknown as Record<string, unknown>)[field] = value;
					slide.elements = [invalid];
					await expect(handler.save(data.slides)).rejects.toThrow(
						new RegExp(`bad-element.*invalid ${field}=`),
					);
					expect(invalid.rawXml).toStrictEqual(baseline);
					expect(slide.id).toBe('slide1');
				}
			}
		},
	);

	it('finds invalid descendants before group child-space inversion', async () => {
		const { handler, data, createSlide } = await PptxHandler.createBlank();
		const child = element();
		child.y = NaN;
		const group = element('group');
		if (group.type !== 'group') throw new Error('Expected group');
		group.children = [{ ...group, id: 'nested', children: [child] }];
		data.slides.push(createSlide().addElement(group).build());
		await expect(handler.save(data.slides)).rejects.toThrow(/bad-element.*invalid y=NaN/);
	});

	it.each([
		'chOffXEmu',
		'chOffYEmu',
		'chExtWidthEmu',
		'chExtHeightEmu',
		'rotation',
		'skewX',
		'skewY',
		'xEmu',
		'yEmu',
		'widthEmu',
		'heightEmu',
	])('rejects non-finite %s before serializing position attributes', (field) => {
		const invalid = element('group');
		(invalid as unknown as Record<string, unknown>)[field] = Infinity;
		expect(() => validateElementGeometry(invalid, 'Slide 2 (slide2)')).toThrow(
			new RegExp(`Slide 2.*${field}=Infinity`),
		);
	});

	it('rejects invalid master and layout elements before save mutation', async () => {
		const { handler, data } = await PptxHandler.createBlank();
		const invalid = element();
		invalid.width = Infinity;
		await expect(
			handler.save(data.slides, {
				slideLayouts: [{ path: 'ppt/slideLayouts/slideLayout1.xml', elements: [invalid] }],
			}),
		).rejects.toThrow(/Template.*slideLayout1.*invalid width=Infinity/);
	});

	it('allows a placeholder with no own frame to keep inheriting absent resolved axes', async () => {
		const { handler, data, createSlide } = await PptxHandler.createBlank();
		const builder = createSlide('Title Slide').addPlaceholder('ctrTitle', 'Inherited');
		const placeholder = builder.getLastElement()!;
		delete (placeholder as unknown as Record<string, unknown>).x;
		data.slides.push(builder.build());
		const zip = await JSZip.loadAsync(await handler.save(data.slides));
		const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
		expect(xml.match(/<p:sp>[\s\S]*?<\/p:sp>/)![0]).not.toContain('<a:xfrm');
	});

	it('rejects an invalid inherited placeholder when a frame override exists', async () => {
		const { handler, data, createSlide } = await PptxHandler.createBlank();
		const builder = createSlide('Title Slide').addPlaceholder('ctrTitle', 'Explicit', { x: 20 });
		builder.getLastElement()!.height = NaN;
		data.slides.push(builder.build());
		await expect(handler.save(data.slides)).rejects.toThrow(/invalid height=NaN/);
	});

	it('rejects raw JavaScript and JSON geometry without coercing strings or null', async () => {
		const { handler, data, createSlide } = await PptxHandler.createBlank();
		for (const x of [null, '100']) {
			const raw = JSON.parse(JSON.stringify({ ...element(), x }));
			data.slides = [createSlide().addElement(raw).build()];
			await expect(handler.save(data.slides)).rejects.toThrow(/invalid x=/);
		}
	});

	it('keeps valid zero and negative offsets and zero-sized connectors unchanged', () => {
		const valid = element('connector');
		Object.assign(valid, { x: -100, y: 0, width: 0, height: 0 });
		expect(() => validateElementGeometry(valid, 'Slide 1')).not.toThrow();
	});

	it('guards the direct transform updater before creating coordinate children', () => {
		const shape = { 'p:spPr': { 'a:xfrm': {} } };
		const invalid = element();
		invalid.x = NaN;
		expect(() => new PptxElementTransformUpdater().applyTransform(shape, invalid, 9525)).toThrow(
			/bad-element.*invalid x=NaN/,
		);
		expect(shape).toStrictEqual({ 'p:spPr': { 'a:xfrm': {} } });
	});
});
