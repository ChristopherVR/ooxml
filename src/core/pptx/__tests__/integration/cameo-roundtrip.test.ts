import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { PresentationBuilder } from '../../core/builders/sdk/PresentationBuilder';
import { PptxHandler } from '../../core/PptxHandler';
import {
	CAMERA_CLIP_EXTENSION,
	containsCameraClip,
	removeCameraClips,
} from '../../editor/render/slide-operations';
import { must } from '../../test-support/access';

const POSTER =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg==';

it('preserves the camera recording marker and leaves ordinary video when resetting it', async () => {
	const created = await PresentationBuilder.create();
	const slide = created.createSlide('Blank').build();
	const video = {
		type: 'media' as const,
		mediaType: 'video' as const,
		mediaData: 'data:video/webm;base64,GkXfow==',
		x: 20,
		y: 20,
		width: 200,
		height: 150,
	};
	slide.elements.push(
		{ ...video, id: 'recorded-camera', extensionXml: [CAMERA_CLIP_EXTENSION] },
		{ ...video, id: 'ordinary-video' },
	);
	created.data.slides.push(slide);
	const handler = new PptxHandler();
	const saved = await handler.load(await created.handler.save(created.data.slides));
	const loaded = must(saved.slides[0]);
	expect(containsCameraClip(loaded.elements)).toBe(true);
	loaded.elements = removeCameraClips(loaded.elements);
	loaded.isDirty = true;
	const reloaded = await new PptxHandler().load(await handler.save(saved.slides));
	const remaining = must(reloaded.slides[0]).elements;
	expect(remaining).toHaveLength(1);
	expect(remaining[0]).toMatchObject({ type: 'media', mediaType: 'video' });
	expect(containsCameraClip(remaining)).toBe(false);
});

it('round-trips a Cameo picture and its poster when the frame is moved', async () => {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	const slide = createSlide('Blank').build();
	slide.elements.push({
		type: 'image',
		id: 'camera',
		name: 'Cameo',
		liveFeed: true,
		cropShape: 'ellipse',
		imageData: POSTER,
		x: 480,
		y: 260,
		width: 144,
		height: 144,
	});
	data.slides.push(slide);
	const bytes = await handler.save(data.slides);
	const loaded = new PptxHandler();
	const deck = await loaded.load(bytes);
	const camera = must(must(deck.slides[0]).elements[0]);
	expect(camera).toMatchObject({ type: 'picture', liveFeed: true, shapeType: 'ellipse', x: 480 });
	camera.x = 420;
	const edited = await loaded.save(deck.slides);
	const zip = await JSZip.loadAsync(edited);
	const xml = await must(zip.file('ppt/slides/slide1.xml')).async('string');
	expect(xml).toContain('alf:liveFeedProps');
	expect(Object.keys(zip.files).some((name) => /^ppt\/media\/.+\.png$/.test(name))).toBe(true);
	const reopened = await new PptxHandler().load(edited);
	expect(reopened.slides[0]?.elements[0]).toMatchObject({ liveFeed: true, x: 420 });
});
