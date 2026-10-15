/**
 * The inspector sections (media, picture, table, cell formatting, text effects,
 * SmartArt and the presentation properties card) name their captions,
 * aria-labels, titles and placeholders through the translator, not in
 * hard-coded English. Mounted with a marking translator every visible string
 * must come back marked; mounted under German, none of the old English
 * literals may reach the DOM and the German words must.
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
import PresentationPropertiesPanel from './PresentationPropertiesPanel.svelte';
import SmartArtSection from './SmartArtSection.svelte';
import TableSection from './TableSection.svelte';
import TextEffectsSection from './TextEffectsSection.svelte';

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

/**
 * Labels the shared catalogues in `ooxml-ui/pptx` hold in English: table style
 * presets, the text warp, 3D material and bevel presets. Every binding (React,
 * Vue, Angular, vanilla and Svelte) paints them unchanged, so they are not
 * Svelte literals. Translating them is a shared change that must reach all
 * five bindings together, and is tracked as a follow-up.
 */
const SHARED_CATALOGUE_LABELS = new Set([
	'Light 1',
	'Light 2',
	'Light 3',
	'Medium 1',
	'Medium 2',
	'Medium 3',
	'Dark 1',
	'Dark 2',
	'Dark 3',
	'Arch Up',
	'Arch Down',
	'Circle',
	'Wave 1',
	'Wave 2',
	'Inflate',
	'Deflate',
	'Fade Right',
	'Slant Up',
	'Triangle',
	'Matte',
	'Plastic',
	'Metal',
	'Warm Matte',
	'Soft Edge',
	'Flat',
	'None',
	'Relaxed Inset',
	'Cool Slant',
	'Angle',
	'Soft Round',
	'Convex',
]);

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

const TABLE = {
	type: 'table',
	id: 'tbl1',
	x: 0,
	y: 0,
	width: 200,
	height: 100,
	tableData: {
		rows: [
			{
				cells: [
					{
						text: 'A',
						style: {
							fillMode: 'gradient',
							gradientFillType: 'linear',
							gradientFillAngle: 90,
							gradientFillStops: [
								{ color: '#000000', position: 0 },
								{ color: '#ffffff', position: 100000 },
							],
						},
					},
					{ text: 'B' },
				],
			},
			{ cells: [{ text: 'C' }, { text: 'D' }] },
		],
		columnWidths: [0.5, 0.5],
		bandedRows: true,
		bandedColumns: true,
	},
} as unknown as PptxElement;

const TEXT = {
	type: 'text',
	id: 't1',
	x: 0,
	y: 0,
	width: 200,
	height: 80,
	textStyle: {
		textShadowColor: '#000000',
		textShadowBlur: 4,
		textShadowOffsetX: 2,
		textShadowOffsetY: 2,
		textShadowOpacity: 0.5,
		textGlowColor: '#ffff00',
		textGlowRadius: 6,
		textGlowOpacity: 0.6,
		textReflection: true,
		textReflectionBlur: 1,
		textReflectionOffset: 3,
		textReflectionStartOpacity: 0.5,
		textReflectionEndOpacity: 0,
		textWarpPreset: 'archUp',
		textWarpAdj: 50000,
		text3d: {
			extrusionHeight: 57150,
			extrusionColor: '#808080',
			presetMaterial: 'plastic',
			bevelTopType: 'circle',
		},
	},
} as unknown as PptxElement;

const SMART_ART = {
	type: 'smartArt',
	id: 'sa1',
	x: 0,
	y: 0,
	width: 400,
	height: 300,
	smartArtData: {
		nodes: [
			{ id: 'n1', text: 'One' },
			{ id: 'n2', text: 'Two' },
		],
		resolvedLayoutType: 'list',
		drawingShapes: [],
	},
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
	return found.filter((text) => !DATA.has(text) && !SHARED_CATALOGUE_LABELS.has(text));
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
	'Cell padding',
	'Row',
	'Column',
	'Insert row',
	'Insert column',
	'Delete row',
	'Delete column',
	'Row band cycle',
	'Column band cycle',
	'Column widths',
	'Row heights',
	'Distribute evenly',
	'Cell formatting',
	'Text color',
	'Background',
	'Font size',
	'Fill',
	'Solid',
	'Pattern',
	'None',
	'Horizontal',
	'Vertical',
	'Type',
	'Angle',
	'Foreground',
	'Stop ',
	'Top margin',
	'Borders',
	'Merge right',
	'Merge down',
	'Split',
	'Text effects',
	'Shadow',
	'Glow',
	'Reflection',
	'Warp',
	'Adjustment 2',
	'3D and bevel',
	'Depth (EMU)',
	'Extrusion',
	'Bevel',
	'Blur',
	'Offset',
	'Start opacity',
	'End opacity',
	'Diagram style',
	'Add',
	'Remove',
	'Up',
	'Down',
	'Promote',
	'Demote',
	'Font',
	'Bold',
	'Italic',
	'Select item',
	'Slide',
	'elements',
];

function mountAll(translator: Translator): HTMLElement[] {
	const media = editorWith(MEDIA);
	const image = editorWith(IMAGE);
	const table = editorWith(TABLE);
	const text = editorWith(TEXT);
	const smartArt = editorWith(SMART_ART);
	// A selected node reveals the Fill, Font, Bold and Italic controls.
	const smartArtTarget = render(SmartArtSection, { editor: smartArt, el: SMART_ART }, translator);
	smartArtTarget.querySelector<HTMLButtonElement>('.pptx-svelte-smartart-nodes button')?.click();
	flushSync();
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
		render(TableSection, { editor: table, el: TABLE }, translator),
		render(TextEffectsSection, { editor: text, el: TEXT }, translator),
		smartArtTarget,
		render(
			PresentationPropertiesPanel,
			{ editor: text, deck: { canvasSize: { width: 960, height: 540 } } },
			translator,
		),
	];
}

describe('inspector sections translation', () => {
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
			'Zeile einfügen',
			'Gleichmäßig verteilen',
			'Zellformatierung',
			'Anpassung 2',
			'Element 1 auswählen',
			'Texteffekte',
			'Rahmenfarbe',
			'Elemente',
		]) {
			expect(painted.some((text) => text.includes(expected))).toBeTruthy();
		}
	});
});
