// @vitest-environment jsdom
import { afterEach, beforeAll, expect, it, vi } from 'vitest';
import type { MediaPptxElement } from 'ooxml-core/pptx';
import './contextual-ribbon-controls';
import { registerPptxWebControls } from './index';

beforeAll(registerPptxWebControls);
afterEach(() => document.body.replaceChildren());

const video: MediaPptxElement = {
	id: 'video',
	type: 'media',
	mediaType: 'video',
	x: 0,
	y: 0,
	width: 160,
	height: 90,
	mediaPath: 'ppt/media/clip.mp4',
	volume: 0.8,
};

it('edits the selected clip and gates unsupported options', async () => {
	const strip = document.createElement('pptx-ui-contextual-ribbon');
	const patch = vi.fn();
	strip.group = 'playback.options';
	strip.state = { element: video, editable: true, translate: (key) => key, patch };
	document.body.append(strip);
	const volume = await vi.waitFor(() => {
		const input = strip.querySelector<HTMLInputElement>(
			'[data-ribbon-control="playback.options.volume"] input',
		);
		expect(input?.value).toBe('80');
		return input!;
	});
	volume.value = '50';
	volume.dispatchEvent(new Event('change'));
	expect(patch).toHaveBeenCalledWith('video', { volume: 0.5 });
	expect(
		strip
			.querySelector('[data-ribbon-control="playback.options.across"]')
			?.hasAttribute('disabled'),
	).toBe(true);
	const start = strip.querySelector<HTMLElementTagNameMap['pptx-ui-select']>(
		'[data-ribbon-control="playback.options.start"] pptx-ui-select',
	)!;
	start.value = 'sequence';
	start.dispatchEvent(new Event('change'));
	expect(patch).toHaveBeenCalledWith('video', {
		startMode: 'sequence',
		autoPlay: false,
		playAcrossSlides: false,
	});
	strip.state = { ...strip.state, editable: false };
	expect(volume.disabled).toBe(true);
});

it('explains why trim and fade are unavailable for media outside a picture', async () => {
	const strip = document.createElement('pptx-ui-contextual-ribbon');
	strip.group = 'playback.editing';
	strip.state = {
		element: { ...video, rawXml: { 'p:nvGraphicFramePr': {} } },
		editable: true,
		translate: (key) => key,
		patch: vi.fn(),
	};
	document.body.append(strip);
	await vi.waitFor(() => {
		const trim = strip.querySelector('[data-ribbon-control="playback.editing.trim"]');
		const fade = strip.querySelector<HTMLInputElement>(
			'[data-ribbon-control="playback.editing.fadeIn"] input',
		);
		expect(trim?.hasAttribute('disabled')).toBe(true);
		expect(fade?.disabled).toBe(true);
		expect(trim?.getAttribute('title')).toBe('This command requires embedded media in a picture.');
		expect(fade?.parentElement?.title).toBe(trim?.getAttribute('title'));
	});
});

it('shows disabled shape commands when the SmartArt data part is missing', async () => {
	const strip = document.createElement('pptx-ui-contextual-ribbon');
	strip.group = 'smartArtFormat.shapes';
	strip.state = {
		element: { id: 'diagram', type: 'smartArt', x: 0, y: 0, width: 160, height: 90 },
		editable: true,
		translate: (key) => key,
		patch: vi.fn(),
	};
	document.body.append(strip);
	await vi.waitFor(() => {
		const larger = strip.querySelector('[data-ribbon-control="smartArtFormat.shapes.larger"]');
		expect(larger?.hasAttribute('disabled')).toBe(true);
		expect(larger?.getAttribute('title')).toBe('The selected node has no editable drawing shape.');
	});
});
