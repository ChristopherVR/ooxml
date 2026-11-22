import { remapZoomTargets } from '../../core/utils/remap-zoom-targets';
import JSZip from 'jszip';
import { describe, expect, it, vi } from 'vitest';

import { NS } from '../../../xml/namespaces';
import { parseXml } from '../../../xml/xml';
import { pptxTools } from '../../automation/catalogue/index';
import { PresentationBuilder } from '../../core/builders/sdk/PresentationBuilder';
import { PptxHandler } from '../../core/PptxHandler';
import type { ZoomPptxElement } from '../../core/types';

const PNG_DATA =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=';

async function buildSlideZoomPackage(): Promise<ArrayBuffer> {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	data.slides.push(createSlide('Blank').build());
	data.slides.push(createSlide('Blank').build());
	data.slides.push(createSlide('Blank').build());
	const zip = await JSZip.loadAsync(await handler.save(data.slides));
	const presentationXml = await zip.file('ppt/presentation.xml')!.async('string');
	const numericIds = [...presentationXml.matchAll(/<p:sldId id="(?<id>\d+)"/gu)].map(
		(match) => match.groups!.id,
	);
	const slidePath = 'ppt/slides/slide1.xml';
	const slideXml = await zip.file(slidePath)!.async('string');
	const zoomXml = `
		<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
			<mc:Choice Requires="pslz" xmlns:pslz="http://schemas.microsoft.com/office/powerpoint/2016/slidezoom" xmlns:p166="http://schemas.microsoft.com/office/powerpoint/2016/6/main">
				<pslz:sldZm>
					<pslz:sldZmObj sldId="${numericIds[2]}" cId="41">
						<pslz:zmPr id="{11111111-2222-3333-4444-555555555555}" imageType="preview" returnToParent="0">
							<p166:blipFill><a:blip r:embed="rIdZoomPreview"/><a:stretch><a:fillRect/></a:stretch></p166:blipFill>
							<p166:spPr><a:xfrm><a:off x="952500" y="476250"/><a:ext cx="1905000" cy="952500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p166:spPr>
						</pslz:zmPr>
						<pslz:extLst><p:ext uri="urn:zoom:preserve"><p:vendorData val="keep"/></p:ext></pslz:extLst>
					</pslz:sldZmObj>
				</pslz:sldZm>
			</mc:Choice>
			<mc:Fallback>
				<p:pic><p:nvPicPr><p:cNvPr id="2" name="Slide Zoom"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdZoomPreview"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="952500" y="476250"/><a:ext cx="1905000" cy="952500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>
			</mc:Fallback>
		</mc:AlternateContent>`;
	zip.file(slidePath, slideXml.replace('</p:spTree>', `${zoomXml}</p:spTree>`));

	const relsPath = 'ppt/slides/_rels/slide1.xml.rels';
	const relsXml = await zip.file(relsPath)!.async('string');
	zip.file(
		relsPath,
		relsXml.replace(
			'</Relationships>',
			'<Relationship Id="rIdZoomPreview" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/zoom-preview.png"/></Relationships>',
		),
	);
	zip.file('ppt/media/zoom-preview.png', Buffer.from(PNG_DATA.split(',')[1], 'base64'));
	const contentTypes = await zip.file('[Content_Types].xml')!.async('string');
	if (!/Extension="png"/u.test(contentTypes)) {
		zip.file(
			'[Content_Types].xml',
			contentTypes.replace(
				'</Types>',
				'<Default Extension="png" ContentType="image/png"/></Types>',
			),
		);
	}
	return (await zip.generateAsync({ type: 'uint8array' })).buffer as ArrayBuffer;
}

describe('slide zoom package integration', () => {
	it('loads actual Zoom AlternateContent and preserves unknown XML on dirty save', async () => {
		const handler = new PptxHandler();
		const data = await handler.load(await buildSlideZoomPackage(), { eagerDecodeImages: true });
		const zoom = data.slides[0].elements.find(
			(element): element is ZoomPptxElement => element.type === 'zoom',
		);

		expect(zoom).toMatchObject({
			zoomType: 'slide',
			targetSlideIndex: 2,
			x: 100,
			y: 50,
			width: 200,
			height: 100,
			imagePath: 'ppt/media/zoom-preview.png',
		});
		expect(zoom?.imageData).toMatch(/^(?:data:image\/png;base64,|blob:)/u);

		zoom!.x = 160;
		zoom!.width = 240;
		zoom!.targetSlideIndex = 1;
		const saved = await handler.save(data.slides);
		const zip = await JSZip.loadAsync(saved);
		const slideXml = await zip.file('ppt/slides/slide1.xml')!.async('string');
		const presentationXml = await zip.file('ppt/presentation.xml')!.async('string');
		const targetId = [...presentationXml.matchAll(/<p:sldId id="(?<id>\d+)"/gu)][1].groups!.id;

		expect(slideXml).toContain('<mc:AlternateContent');
		expect(slideXml).toContain('<pslz:sldZm>');
		expect(slideXml).toContain('<mc:Fallback>');
		expect(slideXml).toContain(`sldId="${targetId}"`);
		expect(slideXml).toContain('x="1524000"');
		expect(slideXml).toContain('cx="2286000"');
		expect(slideXml).toContain('uri="urn:zoom:preserve"');
		expect(slideXml).toContain('val="keep"');
		expect(slideXml).toContain('r:embed="rIdZoomPreview"');

		const reloaded = await new PptxHandler().load(saved.buffer as ArrayBuffer, {
			eagerDecodeImages: true,
		});
		expect(reloaded.slides[0].elements.find((element) => element.type === 'zoom')).toMatchObject({
			type: 'zoom',
			targetSlideIndex: 1,
			x: 160,
			width: 240,
			imagePath: 'ppt/media/zoom-preview.png',
		});
	});

	it('refreshes the native PNG when its target changes, preserving the former image part', async () => {
		const handler = new PptxHandler();
		const reopened = new PptxHandler();
		const updatedPng =
			'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jHuoAAAAASUVORK5CYII=';
		try {
			const { slides } = await handler.load(await buildSlideZoomPackage(), {
				eagerDecodeImages: true,
			});
			const renderZoomPreview = vi.fn(async () => updatedPng);
			await handler.save(slides, { renderZoomPreview });
			expect(renderZoomPreview).not.toHaveBeenCalled();

			slides[2].backgroundColor = '#4e69b3';
			const bytes = await handler.save(slides, { renderZoomPreview });
			expect(renderZoomPreview).toHaveBeenCalledExactlyOnceWith(
				slides[2],
				expect.any(Number),
				expect.any(Number),
				false,
			);
			const saved = await reopened.load(bytes.slice().buffer, { eagerDecodeImages: true });
			const zoom = saved.slides[0].elements.find((element) => element.type === 'zoom');
			expect(zoom?.imagePath).toMatch(/^ppt\/media\/image\d+\.png$/u);
			const zip = await JSZip.loadAsync(bytes);
			expect(await zip.file(zoom!.imagePath!)?.async('base64')).toBe(updatedPng.split(',')[1]);
			expect(await zip.file('ppt/media/zoom-preview.png')?.async('base64')).toBe(
				PNG_DATA.split(',')[1],
			);
			const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
			const slide = parseXml(xml);
			const envelope = slide.getElementsByTagNameNS(NS.mc, 'AlternateContent')[0];
			const blips = Array.from(envelope.getElementsByTagNameNS(NS.a, 'blip'));
			expect(blips).toHaveLength(2);
			const embeds = blips.map((blip) => blip.getAttributeNS(NS.r, 'embed'));
			expect(embeds[0]).toBeTruthy();
			expect(embeds[0]).toBe(embeds[1]);
			expect(embeds[0]).not.toBe('rIdZoomPreview');
		} finally {
			handler.dispose();
			reopened.dispose();
		}
	});

	it('keeps the same Slide Zoom target after reordering and reopening', async () => {
		const handler = new PptxHandler();
		const data = await handler.load(await buildSlideZoomPackage());
		const [first, second, target] = data.slides;
		const reordered = remapZoomTargets(data.slides, [first, target, second]);
		expect(reordered[0].elements.find((element) => element.type === 'zoom')).toMatchObject({
			targetSlideIndex: 1,
		});
		const bytes = await handler.save(reordered);
		const reloaded = await new PptxHandler().load(bytes.buffer as ArrayBuffer);
		expect(reloaded.slides[0].elements.find((element) => element.type === 'zoom')).toMatchObject({
			targetSlideIndex: 1,
		});
		expect(reloaded.slides[1].slideId).toBe(target.slideId);
	});

	it('preserves the Slide Zoom envelope when an automation tool edits speaker notes', async () => {
		const input = new Uint8Array(await buildSlideZoomPackage());
		const tool = pptxTools.find((candidate) => candidate.name === 'update_slide_properties');
		if (tool?.kind !== 'edit') throw new Error('Missing update_slide_properties tool');
		const notes = 'Describe the menu before following the Slide Zoom.';
		const { bytes } = await tool.run(input, { slideIndex: 0, notes });
		if (!bytes) throw new Error('The tool did not save the deck');

		const zip = await JSZip.loadAsync(bytes);
		const slideXml = await zip.file('ppt/slides/slide1.xml')!.async('string');
		const slide = parseXml(slideXml);
		const envelope = slide.getElementsByTagNameNS(NS.mc, 'AlternateContent')[0];
		const choice = envelope?.getElementsByTagNameNS(NS.mc, 'Choice')[0];
		const fallback = envelope?.getElementsByTagNameNS(NS.mc, 'Fallback')[0];
		expect(choice?.getAttribute('Requires')).toBe('pslz');
		expect(choice?.getElementsByTagNameNS(NS.p, 'vendorData')[0]?.getAttribute('val')).toBe('keep');
		expect(fallback?.getElementsByTagNameNS(NS.p, 'pic')[0]).toBeDefined();

		const handler = new PptxHandler();
		try {
			const reloaded = await handler.load(new Uint8Array(bytes).buffer);
			expect(reloaded.slides[0].notes).toBe(notes);
			expect(reloaded.slides[0].elements.find((element) => element.type === 'zoom')).toMatchObject({
				zoomType: 'slide',
				targetSlideIndex: 2,
				imagePath: 'ppt/media/zoom-preview.png',
			});
		} finally {
			handler.dispose();
		}
	});

	it('serializes a slide Zoom cover and background options with its image relationship', async () => {
		const { handler, data, createSlide } = await PresentationBuilder.create();
		data.slides.push(createSlide('Blank').build());
		data.slides.push(createSlide('Blank').build());
		data.slides[0].elements.push({
			id: 'authored-slide-zoom',
			type: 'zoom',
			zoomType: 'slide',
			targetSlideIndex: 1,
			zoomImageType: 'cover',
			zoomShowBackground: false,
			x: 80,
			y: 90,
			width: 320,
			height: 180,
			imageData: PNG_DATA,
			altText: 'Preview of slide 2',
			returnToParent: false,
			transitionDurationMs: 750,
		} satisfies ZoomPptxElement);

		const saved = await handler.save(data.slides);
		const zip = await JSZip.loadAsync(saved);
		const slideXml = await zip.file('ppt/slides/slide1.xml')!.async('string');
		const slideRels = await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string');
		const contentTypes = await zip.file('[Content_Types].xml')!.async('string');

		expect(slideXml).toContain('Requires="pslz"');
		const parsedSlide = parseXml(slideXml);
		const zoomData = Array.from(parsedSlide.getElementsByTagNameNS(NS.a, 'graphicData')).find(
			(graphicData) =>
				graphicData.getAttribute('uri') ===
				'http://schemas.microsoft.com/office/powerpoint/2016/slidezoom',
		);
		const frame = zoomData?.parentNode?.parentNode;
		expect(frame?.nodeName).toBe('p:graphicFrame');
		expect(Array.from(frame?.childNodes ?? [], (node) => node.nodeName)).toEqual([
			'p:nvGraphicFramePr',
			'p:xfrm',
			'a:graphic',
		]);
		expect(slideXml).toContain('returnToParent="0"');
		expect(slideXml).toContain('transitionDur="750"');
		expect(slideXml).toContain('imageType="cover"');
		expect(slideXml).toContain('showBg="0"');
		expect(slideXml).toContain('<pslz:sldZm>');
		expect(slideXml).toContain(
			'<p:cNvPr id="2" name="authored-slide-zoom" descr="Preview of slide 2"',
		);
		expect(slideRels).toContain('/relationships/image');
		expect(slideRels).toMatch(/Target="\.\.\/media\/image\d+\.png"/u);
		expect(contentTypes).toContain('Extension="png" ContentType="image/png"');

		const reloaded = await new PptxHandler().load(saved.buffer as ArrayBuffer, {
			eagerDecodeImages: true,
		});
		expect(reloaded.slides[0].elements.find((element) => element.type === 'zoom')).toMatchObject({
			type: 'zoom',
			targetSlideIndex: 1,
			zoomImageType: 'cover',
			zoomShowBackground: false,
			x: 80,
			y: 90,
			width: 320,
			height: 180,
		});
	});
});
