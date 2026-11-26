import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { PptxHandler } from '../../core/PptxHandler';
import { PresentationBuilder } from '../../core/builders/sdk/PresentationBuilder';
import { TRANSPARENT_PNG_DATA_URL } from '../../core/core/runtime/picture-svg-blip';
import { at, must } from '../../test-support/access';

it('persists first-save media playback and picture edits, then clears an edited extension', async () => {
	const { handler, data, createSlide } = await PresentationBuilder.create({ initialSlideCount: 0 });
	const slide = createSlide('Blank')
		.addMedia('video', 'data:video/mp4;base64,AAAA', {
			x: 20,
			y: 30,
			width: 320,
			height: 240,
		})
		.build();
	const video = must(slide.elements.find((element) => element.type === 'media'));
	Object.assign(video, {
		posterFrameData: TRANSPARENT_PNG_DATA_URL,
		shapeType: 'ellipse',
		shapeStyle: { strokeColor: '#112233', strokeFillMode: 'solid', strokeWidth: 2 },
		imageEffects: { brightness: 15 },
		cropLeft: 0.2,
		trimStartMs: 200,
		trimEndMs: 800,
		fadeInDuration: 0.3,
		fadeOutDuration: 0.4,
		bookmarks: [{ id: 'first', time: 0.5, label: 'Bookmark 1' }],
		volume: 0.5,
		autoPlay: true,
		startMode: 'auto',
		rewindAfterPlaying: true,
		loop: true,
		fullScreen: true,
		hideWhenNotPlaying: true,
	});
	data.slides.push(slide);
	const saved = await handler.save(data.slides);
	const zip = await JSZip.loadAsync(saved);
	const relationships = await must(zip.file('ppt/slides/_rels/slide1.xml.rels')).async('text');
	expect(relationships).toContain('http://schemas.microsoft.com/office/2007/relationships/media');
	const xml = await must(zip.file('ppt/slides/slide1.xml')).async('text');
	expect(xml).toMatch(/<p14:media[^>]+r:embed="rId\d+"[^>]*>.*<p14:bmkLst>/su);
	expect(xml).toContain('<p14:trim st="200" end="800"');
	expect(xml).toContain('<p14:fade in="300" out="400"');
	expect(xml).toContain('<p:cMediaNode');
	expect(xml).toContain('cmd="playFrom(0.0)"');
	const reload = new PptxHandler();
	const edited = await reload.load(saved.slice().buffer);
	const clip = must(at(edited.slides, 0).elements.find((element) => element.type === 'media'));
	expect(clip).toMatchObject({
		posterFramePath: expect.stringContaining('ppt/media/'),
		shapeType: 'ellipse',
		imageEffects: { brightness: 15 },
		shapeStyle: { strokeColor: '#112233', strokeWidth: 2 },
		cropLeft: 0.2,
		trimStartMs: 200,
		trimEndMs: 800,
		fadeInDuration: 0.3,
		fadeOutDuration: 0.4,
		bookmarks: [{ time: 0.5, label: 'Bookmark 1' }],
		volume: 0.5,
		autoPlay: true,
		startMode: 'auto',
		rewindAfterPlaying: true,
		loop: true,
		fullScreen: true,
		hideWhenNotPlaying: true,
	});
	clip.startMode = 'sequence';
	clip.autoPlay = false;
	const queued = await reload.save(edited.slides);
	const queueReader = new PptxHandler();
	const queueDeck = await queueReader.load(queued.slice().buffer);
	const queuedClip = must(
		at(queueDeck.slides, 0).elements.find((element) => element.type === 'media'),
	);
	expect(queuedClip.startMode).toBe('sequence');
	Object.assign(queuedClip, {
		bookmarks: [],
		trimStartMs: undefined,
		trimEndMs: undefined,
		fadeInDuration: undefined,
		fadeOutDuration: undefined,
		autoPlay: false,
		startMode: 'click',
		rewindAfterPlaying: false,
		loop: false,
		fullScreen: false,
		hideWhenNotPlaying: false,
	});
	const cleared = await queueReader.save(queueDeck.slides);
	const clearedZip = await JSZip.loadAsync(cleared);
	const clearedXml = await must(clearedZip.file('ppt/slides/slide1.xml')).async('text');
	expect(clearedXml).not.toMatch(/p14:(?:bmkLst|trim|fade)/u);
	expect(clearedXml).not.toContain('cmd="playFrom(0.0)"');
	const reloaded = await new PptxHandler().load(cleared.slice().buffer);
	const replaced = must(
		at(reloaded.slides, 0).elements.find((element) => element.type === 'media'),
	);
	expect(replaced.volume).toBe(0.5);
	for (const key of [
		'autoPlay',
		'loop',
		'fullScreen',
		'hideWhenNotPlaying',
		'rewindAfterPlaying',
	] as const)
		expect(replaced[key]).toBeUndefined();
});

it('preserves extra media start conditions on an unedited save', async () => {
	const { handler, data, createSlide } = await PresentationBuilder.create({ initialSlideCount: 0 });
	const slide = createSlide('Blank')
		.addMedia('video', 'data:video/mp4;base64,AAAA', { x: 0, y: 0, width: 100, height: 100 })
		.build();
	must(slide.elements.find((element) => element.type === 'media')).startMode = 'sequence';
	data.slides.push(slide);
	const path = 'ppt/slides/slide1.xml';
	const zip = await JSZip.loadAsync(await handler.save(data.slides));
	const xml = await must(zip.file(path)).async('text');
	const condition = '<p:cond evt="onBegin" delay="150"><p:tn val="1"/></p:cond>';
	zip.file(path, xml.replace(/(<p:cMediaNode[^>]*><p:cTn[^>]*><p:stCondLst>)/u, '$1' + condition));
	const reader = new PptxHandler();
	const loaded = await reader.load(await zip.generateAsync({ type: 'arraybuffer' }));
	const clip = must(at(loaded.slides, 0).elements.find((element) => element.type === 'media'));
	expect(clip.startMode).toBe('sequence');
	const saved = await JSZip.loadAsync(await reader.save(loaded.slides));
	expect(await must(saved.file(path)).async('text')).toContain(condition);
});
