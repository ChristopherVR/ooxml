import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { PptxHandler } from '../../PptxHandler';
import { xmlPath } from '../../utils/xml-access';
import { Presentation } from './Presentation';
import { PresentationBuilder } from './PresentationBuilder';
import { SlideBuilder } from './SlideBuilder';
import { duplicateSlide } from './slide-operations';

function textElement(element: ReturnType<SlideBuilder['build']>['elements'][number]) {
	if (element.type !== 'text' && element.type !== 'shape') throw new Error('Expected text');
	return element;
}

async function slideXml(bytes: Uint8Array, index = 1) {
	return (await JSZip.loadAsync(bytes)).file(`ppt/slides/slide${index}.xml`)!.async('string');
}

describe('SlideBuilder.addPlaceholder', () => {
	it.each([
		['Title Slide', 'ctrTitle'],
		['Title Slide', 'subTitle'],
		['Title and Content', 'title'],
		['Title and Content', 'body'],
		['Section Header', 'body'],
		['Two Content', 'body'],
		['Title Only', 'dt'],
		['Title Only', 'ftr'],
		['Title Only', 'sldNum'],
	] as const)('inherits %s / %s without authoring overrides', async (layout, type) => {
		const { handler, data, createSlide } = await PresentationBuilder.create();
		const builder = createSlide(layout).addPlaceholder(type, 'Inherited text');
		const element = textElement(builder.getLastElement()!);
		expect(element.inheritedTransform).toBeDefined();
		expect(element.width).toBeGreaterThan(0);
		expect(element.height).toBeGreaterThan(0);
		expect(element.textStyle?.fontSize).toBeGreaterThan(0);
		data.slides.push(builder.build());
		const bytes = await handler.save(data.slides);
		const xml = await slideXml(bytes);
		expect(xml).toContain('<p:ph');
		const shapeXml = xml.match(/<p:sp>[\s\S]*?<\/p:sp>/)![0];
		expect(shapeXml).not.toContain('<a:xfrm');
		expect(shapeXml).not.toMatch(/\b(sz|lIns|rIns|tIns|bIns)=/);
		expect(shapeXml).not.toContain('<a:solidFill');
		const reloaded = await new PptxHandler().load(bytes.buffer as ArrayBuffer);
		const actual = textElement(reloaded.slides[0].elements[0]);
		expect(actual.placeholderType).toBe(element.placeholderType);
		for (const key of ['x', 'y', 'width', 'height'] as const)
			expect(actual[key]).toBe(element[key]);
		expect(actual.textStyle?.fontSize).toBe(element.textStyle?.fontSize);
		expect(actual.textStyle?.fontFamily).toBe(element.textStyle?.fontFamily);
		expect(actual.textStyle?.color).toBe(element.textStyle?.color);
	});

	it('inherits the layout language instead of stamping a default one', async () => {
		const created = await PresentationBuilder.create();
		created.data.slides.push(created.createSlide('Title Slide').build());
		const zip = await JSZip.loadAsync(await created.handler.save(created.data.slides));
		const layoutPath = 'ppt/slideLayouts/slideLayout1.xml';
		const layoutXml = (await zip.file(layoutPath)!.async('string')).replace(
			/<p:sp\b[\s\S]*?<\/p:sp>/g,
			(shape) =>
				shape.includes('type="ctrTitle"')
					? shape.replace(
							/<a:lstStyle\s*(?:\/>|>[\s\S]*?<\/a:lstStyle>)/,
							'<a:lstStyle><a:lvl1pPr><a:defRPr lang="fr-FR"/></a:lvl1pPr></a:lstStyle>',
						)
					: shape,
		);
		expect(layoutXml).toContain('lang="fr-FR"');
		zip.file(layoutPath, layoutXml);
		const template = await zip.generateAsync({ type: 'uint8array' });
		const handler = new PptxHandler();
		await handler.load(template.buffer as ArrayBuffer);
		const slide = new SlideBuilder(1, layoutPath, 'Title Slide', handler)
			.addPlaceholder('ctrTitle', 'Bonjour')
			.addPlaceholder('subTitle', [{ text: 'Gras', style: { bold: true } }])
			.build();
		const bytes = await handler.save([slide]);
		const saved = await JSZip.loadAsync(bytes);
		const slidePart = Object.keys(saved.files).find((name) =>
			/^ppt\/slides\/slide\d+\.xml$/.test(name),
		)!;
		const xml = await saved.file(slidePart)!.async('string');
		expect(xml).toContain('Bonjour');
		expect(xml).not.toContain('lang="en-US"');
		expect(xml).toContain('b="1"');
		const reloaded = await new PptxHandler().load(bytes.buffer as ArrayBuffer);
		expect(reloaded.slides[0].elements.map((element) => element.placeholderType)).toEqual([
			'ctrtitle',
			'subtitle',
		]);
	});

	it.each([
		['empty text', '', undefined],
		['a newline only', '\n', undefined],
		['an empty bold segment', [{ text: '', style: { bold: true } }], 'b="1"'],
	] as const)('keeps %s free of a default language override', async (_label, text, expected) => {
		const { handler, createSlide } = await PresentationBuilder.create();
		const saved = await JSZip.loadAsync(
			await handler.save([
				createSlide('Title Slide')
					.addPlaceholder('ctrTitle', text as string | { text: string; style: { bold: true } }[])
					.build(),
			]),
		);
		const slidePart = Object.keys(saved.files).find((name) =>
			/^ppt\/slides\/slide\d+\.xml$/.test(name),
		)!;
		const xml = await saved.file(slidePart)!.async('string');
		expect(xml).toContain('<p:ph type="ctrTitle"');
		expect(xml).not.toContain('lang="en-US"');
		if (expected) expect(xml).toContain(expected);
	});

	it('targets both content indices and preserves explicit rich text only', async () => {
		const { handler, data, createSlide } = await PresentationBuilder.create();
		const builder = createSlide('Two Content')
			.addPlaceholder('body', 'Left', { idx: 1 })
			.addPlaceholder('body', [{ text: 'Right', style: { bold: true } }], { idx: 2 });
		const [left, right] = builder.getElements().map(textElement);
		expect(left.x).toBeLessThan(right.x);
		expect(xmlPath(right.rawXml, 'p:nvSpPr', 'p:nvPr', 'p:ph')?.['@_idx']).toBe('2');
		data.slides.push(builder.build());
		const xml = await slideXml(await handler.save(data.slides));
		expect(xml).toContain('idx="2"');
		expect(xml).toContain('b="1"');
		expect(xml).not.toMatch(/sz="\d+"/);
	});

	it('keeps mixed SDK shape ids unique and stable through two saves and reloads', async () => {
		const { handler, data, createSlide } = await PresentationBuilder.create();
		data.slides.push(
			createSlide('Two Content')
				.addPlaceholder('title', 'Heading')
				.addPlaceholder('body', 'Left', { idx: 1 })
				.addPlaceholder('body', 'Right', { idx: 2 })
				.addText('Free text', { x: 10, y: 10, width: 100, height: 20 })
				.addShape('rect', { x: 20, y: 50, width: 100, height: 40 })
				.build(),
		);
		const bytes = await handler.save(data.slides);
		const ids = (xml: string) =>
			[...xml.matchAll(/<p:cNvPr\b[^>]*\bid="(\d+)"/g)].map((match) => match[1]);
		const first = ids(await slideXml(bytes));
		expect(first).toHaveLength(6);
		expect(new Set(first).size).toBe(first.length);
		const loadedHandler = new PptxHandler();
		const loaded = await loadedHandler.load(bytes.buffer as ArrayBuffer);
		expect(
			loaded.slides[0].elements.map(
				(element) => xmlPath(element.rawXml, 'p:nvSpPr', 'p:cNvPr')?.['@_id'],
			),
		).toEqual(first.slice(1));
		loaded.slides[0].isDirty = true;
		const again = await loadedHandler.save(loaded.slides);
		expect(ids(await slideXml(again))).toEqual(first);
		const reloaded = await new PptxHandler().load(again.buffer as ArrayBuffer);
		expect(
			reloaded.slides[0].elements.map(
				(element) => xmlPath(element.rawXml, 'p:nvSpPr', 'p:cNvPr')?.['@_id'],
			),
		).toEqual(first.slice(1));
	});

	it('keeps Blank empty and gives every named starter layout real slots', async () => {
		const { handler, data } = await PresentationBuilder.create();
		const zip = await JSZip.loadAsync(await handler.save(data.slides));
		for (const layout of handler.getLayoutOptions()) {
			const xml = await zip.file(layout.path)!.async('string');
			if (layout.name === 'Blank') expect(xml).not.toContain('<p:ph');
			else expect(xml).toContain('<p:ph');
		}
	});

	it('materializes a transform only for an explicit frame or a viewer-style move', async () => {
		const { handler, data, createSlide } = await PresentationBuilder.create();
		const builder = createSlide('Title Slide').addPlaceholder('ctrTitle', 'Moved');
		const element = builder.getLastElement()!;
		element.x += 20;
		data.slides.push(builder.build());
		const xml = await slideXml(await handler.save(data.slides));
		expect(xml.match(/<p:sp>[\s\S]*?<\/p:sp>/)![0]).toContain('<a:xfrm');
		const explicit = createSlide('Title Slide').addPlaceholder('ctrTitle', 'Override', { x: 123 });
		expect(explicit.getLastElement()!.x).toBe(123);
		expect(explicit.getLastElement()!.inheritedTransform).toBeUndefined();
	});

	it('duplicates inherited placeholders without pinning their transform', async () => {
		const { handler, data, createSlide } = await PresentationBuilder.create();
		data.slides.push(createSlide('Title Slide').addPlaceholder('ctrTitle', 'Original').build());
		data.slides.push(duplicateSlide(data.slides[0], 2));
		const xml = await slideXml(await handler.save(data.slides), 2);
		expect(xml.match(/<p:sp>[\s\S]*?<\/p:sp>/)![0]).not.toContain('<a:xfrm');
		expect(xml).toContain('Original');
	});

	it('reports missing context, missing slots, invalid indices, and invalid frame values', async () => {
		expect(() => new SlideBuilder(3).addPlaceholder('title', 'No context')).toThrow(
			/Slide 3.*title.*Title Slide.*loaded handler/,
		);
		const { createSlide } = await PresentationBuilder.create();
		expect(() => createSlide('Blank').addPlaceholder('title', 'No slot')).toThrow(
			/Slide 1.*title.*Blank.*No matching placeholder/,
		);
		expect(() => createSlide('Two Content').addPlaceholder('body', 'No idx', { idx: 99 })).toThrow(
			/idx=99.*Two Content/,
		);
		expect(() => createSlide('Title Slide').addPlaceholder('ctrTitle', 'Bad', { x: NaN })).toThrow(
			/frame for x/,
		);
		expect(() => createSlide('Title Slide').addPlaceholder('ctrTitle', 'Bad', { idx: -1 })).toThrow(
			/unsigned/,
		);
		const explicit = createSlide('Blank').addPlaceholder('title', 'Explicit', {
			x: 1,
			y: 2,
			width: 300,
			height: 40,
		});
		expect(explicit.getLastElement()!.x).toBe(1);
	});

	it('works through the high-level API for created and loaded decks', async () => {
		const presentation = await Presentation.create();
		presentation.addSlide('Title Slide').addPlaceholder('ctrTitle', 'Created');
		const loaded = await Presentation.load((await presentation.save()).buffer as ArrayBuffer);
		loaded.addSlide('Two Content').addPlaceholder('body', 'Loaded', { idx: 2 });
		const again = await new PptxHandler().load((await loaded.save()).buffer as ArrayBuffer);
		expect(again.slides).toHaveLength(2);
		expect(textElement(again.slides[1].elements[0]).text).toContain('Loaded');
	});

	it('resolves geometry through the master when the layout has no transform', async () => {
		const { handler, data, createSlide } = await PresentationBuilder.create();
		data.slides.push(createSlide('Title Slide').build());
		const zip = await JSZip.loadAsync(await handler.save(data.slides));
		const layoutPath = 'ppt/slideLayouts/slideLayout1.xml';
		const layout = await zip.file(layoutPath)!.async('string');
		const title = layout.match(/<p:sp>[\s\S]*?<\/p:sp>/)![0];
		zip.file(layoutPath, layout.replace(title, title.replace(/<a:xfrm>[\s\S]*?<\/a:xfrm>/, '')));
		const masterPath = 'ppt/slideMasters/slideMaster1.xml';
		zip.file(
			masterPath,
			(await zip.file(masterPath)!.async('string')).replace('</p:spTree>', `${title}</p:spTree>`),
		);
		const loadedHandler = new PptxHandler();
		await loadedHandler.load(await zip.generateAsync({ type: 'arraybuffer' }));
		const element = new SlideBuilder(2, layoutPath, 'Title Slide', loadedHandler)
			.addPlaceholder('ctrTitle', 'Master frame')
			.getLastElement()!;
		expect(element.width).toBeGreaterThan(100);
		expect(element.inheritedTransform).toBeDefined();
	});

	it('retains content and a matching frame through a layout switch and reload', async () => {
		const { handler, data, createSlide } = await PresentationBuilder.create();
		data.slides.push(
			createSlide('Title Slide')
				.addPlaceholder('ctrTitle', 'Heading')
				.addPlaceholder('subTitle', 'Content')
				.build(),
		);
		const target = handler.getLayoutOptions().find((layout) => layout.name === 'Two Content')!;
		const updated = await handler.applyLayoutToSlide(0, target.path, data.slides);
		const heading = updated.elements.find(
			(element) => 'text' in element && element.text === 'Heading',
		)!;
		expect(heading.placeholderType).toBe('title');
		const reloaded = await new PptxHandler().load(
			(await handler.save(data.slides)).buffer as ArrayBuffer,
		);
		const actual = reloaded.slides[0].elements.find(
			(element) => 'text' in element && element.text?.includes('Heading'),
		)!;
		expect(actual.x).toBe(heading.x);
		expect(actual.width).toBe(heading.width);
		expect(actual.placeholderType).toBe('title');
	});

	it('preserves unsaved content and formatting through the existing layout switch path', async () => {
		const bytes = await readFile(
			new URL(
				'../../../__tests__/fixtures/corpus/master-layout-inheritance-fills.pptx',
				import.meta.url,
			),
		);
		const handler = new PptxHandler();
		await handler.load(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength));
		const layouts = handler.getLayoutOptions();
		const source = layouts.find((layout) => layout.name === 'Title Slide')!;
		const target = layouts.find((layout) => layout.name === 'Section Header')!;
		const slide = new SlideBuilder(1, source.path, source.name, handler)
			.addPlaceholder('subTitle', 'Original')
			.build();
		const element = textElement(slide.elements[0]);
		element.text = 'Unsaved content';
		element.textSegments![0].text = 'Unsaved content';
		element.textStyle!.bold = true;
		element.textStyle!.hyperlink = 'https://example.com';
		const slides = [slide];
		const updated = await handler.applyLayoutToSlide(0, target.path, slides);
		const current = textElement(updated.elements.find((candidate) => candidate.id === element.id)!);
		expect(current.text).toContain('Unsaved content');
		expect(current.textStyle!.bold).toBe(true);
		const reloaded = await new PptxHandler().load(
			(await handler.save(slides)).buffer as ArrayBuffer,
		);
		const after = textElement(
			reloaded.slides[0].elements.find(
				(candidate) => 'text' in candidate && candidate.text?.includes('Unsaved content'),
			)!,
		);
		expect(after.textStyle!.bold).toBe(true);
		expect(
			after.textSegments!.some((segment) => segment.style?.hyperlink === 'https://example.com'),
		).toBe(true);
	});

	it.each(['Title Slide', 'Title and Content', 'Two Content', 'Section Header'])(
		'resolves a real template %s using the loader cascade',
		async (name) => {
			const bytes = await readFile(
				new URL(
					'../../../__tests__/fixtures/corpus/master-layout-inheritance-fills.pptx',
					import.meta.url,
				),
			);
			const handler = new PptxHandler();
			const data = await handler.load(
				bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
			);
			const layout = handler.getLayoutOptions().find((item) => item.name === name)!;
			const builder = new SlideBuilder(
				data.slides.length + 1,
				layout.path,
				layout.name,
				handler,
			).addPlaceholder(name === 'Title Slide' ? 'subTitle' : 'body', 'Template content');
			const before = textElement(builder.getLastElement()!);
			if (name === 'Title and Content') {
				// This template's obj slot inherits a differently indexed body
				// frame. Only SDK creation resolves it, with a local transform.
				expect(before.inheritedTransform).toBeUndefined();
				expect(xmlPath(before.rawXml, 'p:spPr', 'a:xfrm')).toBeDefined();
				expect(before.width).toBeGreaterThan(100);
			}
			data.slides.push(builder.build());
			const reloaded = await new PptxHandler().load(
				(await handler.save(data.slides)).buffer as ArrayBuffer,
			);
			const after = textElement(reloaded.slides.at(-1)!.elements[0]);
			expect(after.inheritedTransform).toEqual(before.inheritedTransform);
			expect(after.x).toBe(before.x);
			expect(after.textStyle?.fontSize).toBe(before.textStyle?.fontSize);
		},
	);
});
