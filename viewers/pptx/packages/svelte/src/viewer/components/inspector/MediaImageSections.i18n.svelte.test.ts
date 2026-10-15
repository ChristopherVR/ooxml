/**
 * The media inspector, the picture inspector and the trim timeline name their
 * captions, aria-labels, titles and placeholders through the translator, not in
 * hard-coded English. Mounted with a marking translator every visible string
 * must come back marked; mounted under German, none of the old English literals
 * may reach the DOM and the German words must.
 */
import type { MediaPptxElement, PptxElement } from 'pptx-viewer-core';
import { translationsDe } from 'pptx-viewer-locales';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it } from 'vitest';

import { I18N_CONTEXT_KEY } from '../../../i18n/context';
import { interpolate, translate } from '../../../i18n/translator';
import type { Translator } from '../../../i18n/translator';
import { EditorState } from '../../editor/editor-state.svelte';
import ImageSection from './ImageSection.svelte';
import MediaSection from './MediaSection.svelte';
import MediaTrimTimeline from './MediaTrimTimeline.svelte';

let cleanups: Array<() => void> = [];
afterEach(() => {
	for (const cleanup of cleanups) {
		cleanup();
	}
	cleanups = [];
});

const mark: Translator = (key) => `<${key}>`;
const german: Translator = (key, params) => {
	const message = (translationsDe as Record<string, string>)[key];
	return message === undefined ? translate('en', key, params) : interpolate(message, params);
};

/** Data the components render from the document itself (names, not captions). */
const DATA = new Set(['Commentary', 'Intro']);

const MEDIA: MediaPptxElement = {
	type: 'media',
	id: 'm1',
	x: 0,
	y: 0,
	width: 100,
	height: 60,
	mediaType: 'video',
	mediaPath: 'media1.mp4',
	mediaData: 'data:video/mp4;base64,AAAA',
	trimStartMs: 0,
	trimEndMs: 0,
	volume: 1,
	playbackSpeed: 1,
	fadeInDuration: 0,
	fadeOutDuration: 0,
	bookmarks: [{ id: 'b1', label: 'Intro', time: 2 }],
	captionTracks: [
		{ id: 'c1', label: 'Commentary', language: 'en', kind: 'captions', isDefault: true, src: '' },
	],
} as unknown as MediaPptxElement;

const IMAGE: PptxElement = {
	type: 'image',
	id: 'img1',
	x: 0,
	y: 0,
	width: 100,
	height: 100,
	imagePath: 'ppt/media/image1.png',
	imageEffects: { colorWash: { color: '#0066cc', opacity: 40 } },
} as unknown as PptxElement;

function editorWith(element: PptxElement): EditorState {
	const editor = new EditorState({ getCurrent: () => 0, getHandler: () => null });
	editor.editable = true;
	editor.setSlides([{ id: 's1', rId: 'rId1', slideNumber: 1, elements: [element] }]);
	editor.select(element.id);
	return editor;
}

/** Mount a component into a fresh container under the given translator. */
function render(
	component: Parameters<typeof mount>[0],
	props: Record<string, unknown>,
	translator: Translator,
): HTMLElement {
	const target = document.createElement('div');
	document.body.appendChild(target);
	const instance = mount(component, {
		target,
		props,
		context: new Map([[I18N_CONTEXT_KEY, translator]]),
	});
	flushSync();
	cleanups.push(() => {
		void unmount(instance);
		target.remove();
	});
	return target;
}

/** Every visible caption, aria-label, title, placeholder and track label the sections paint. */
function visibleStrings(target: HTMLElement): string[] {
	const found: string[] = [];
	const walker = document.createTreeWalker(target, NodeFilter.SHOW_TEXT);
	for (let node = walker.nextNode(); node; node = walker.nextNode()) {
		const text = node.textContent?.trim() ?? '';
		if (/[A-Za-z]{2,}/u.test(text)) {
			found.push(text);
		}
	}
	for (const element of target.querySelectorAll(
		'[aria-label], [placeholder], [title], track[label]',
	)) {
		for (const name of ['aria-label', 'placeholder', 'title', 'label']) {
			const value = element.getAttribute(name);
			if (value && /[A-Za-z]{2,}/u.test(value)) {
				found.push(value);
			}
		}
	}
	return found.filter((text) => !DATA.has(text));
}

/** The English literals these sections used to hard-code, minus the words German shares. */
const ENGLISH_LITERALS = [
	'Auto play',
	'Loop',
	'Full screen',
	'Across slides',
	'Hide when stopped',
	'Volume',
	'Speed',
	'Trim start ms',
	'Trim end ms',
	'Fade in seconds',
	'Fade out seconds',
	'Bookmarks',
	'Remove bookmark',
	'Bookmark label',
	'Add bookmark',
	'Caption tracks',
	'Label',
	'Language',
	'Kind',
	'Default',
	'Source',
	'Inline WebVTT',
	'Remove track',
	'Add caption track',
	'Captions',
	'Color wash',
	'Wash color',
	'Wash opacity',
	'Transparency',
	'Bi-level threshold',
	'Play',
	'Track',
	'Media trim timeline',
	'Trim start',
	'Trim end',
	'captions.vtt or data URL',
];

function mountAll(translator: Translator): HTMLElement[] {
	const media = editorWith(MEDIA);
	const image = editorWith(IMAGE);
	return [
		render(MediaSection, { editor: media }, translator),
		render(ImageSection, { editor: image, el: IMAGE }, translator),
		render(
			MediaTrimTimeline,
			{
				duration: 12,
				startMs: 0,
				endMs: 0,
				currentTime: 0,
				bookmarks: MEDIA.bookmarks ?? [],
				onchange: () => {},
				onseek: () => {},
			},
			translator,
		),
	];
}

describe('media and image inspectors translation', () => {
	it('spells every caption, aria-label, title and placeholder through the translator', () => {
		const english = mountAll(mark).flatMap((target) => visibleStrings(target));

		expect(english.filter((text) => !text.includes('<pptx.'))).toStrictEqual([]);
	});

	it('paints German strings and none of the old English literals under a German dictionary', () => {
		const painted = mountAll(german).flatMap((target) => visibleStrings(target));

		for (const literal of ENGLISH_LITERALS) {
			expect(painted.filter((text) => text.includes(literal))).toStrictEqual([]);
		}
		for (const expected of [
			'Automatisch wiedergeben',
			'Endlosschleife',
			'Vollbild',
			'Lautstärke',
			'Untertitelspuren',
			'Lesezeichen hinzufügen',
			'Farbüberlagerung',
			'Opazität',
			'Transparenz',
			'Spur entfernen',
		]) {
			expect(painted.some((text) => text.includes(expected))).toBeTruthy();
		}
	});
});
