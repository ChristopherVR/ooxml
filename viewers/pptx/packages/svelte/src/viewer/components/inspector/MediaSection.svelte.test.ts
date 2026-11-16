/**
 * MediaSection.svelte: the "Trim end ms" field used to bind `trimEndMs`
 * directly.
 *
 * `trimEndMs` is `p14:trim/@end`'s distance from the clip's TAIL
 * (COM-verified, see shared `media-trim-range.ts`), not an absolute stop
 * time. React's `MediaInspector` and Vue's `MediaPropertiesPanel.vue` show an
 * absolute end position and convert back to the tail distance on commit; this
 * field skipped that conversion, so typing "the last 5s" of a 20s clip meant
 * computing 20000-5000 by hand instead of typing 15000.
 */
import type { MediaPptxElement, PptxElement, PptxSlide } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { EditorState } from '../../editor/editor-state.svelte';
import MediaSection from './MediaSection.svelte';

let cleanup: (() => void) | undefined;

afterEach(() => {
	cleanup?.();
	cleanup = undefined;
	vi.restoreAllMocks();
});

function mediaElement(overrides: Partial<MediaPptxElement> = {}): MediaPptxElement {
	return {
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
		...overrides,
	} as MediaPptxElement;
}

function makeEditor(element: MediaPptxElement): EditorState {
	const editor = new EditorState({ getCurrent: () => 0, getHandler: () => null });
	editor.editable = true;
	const slide: PptxSlide = {
		id: 's1',
		rId: 'rId1',
		slideNumber: 1,
		elements: [element as unknown as PptxElement],
	};
	editor.setSlides([slide]);
	editor.select(element.id);
	return editor;
}

function mountSection(editor: EditorState): HTMLElement {
	const target = document.createElement('div');
	document.body.appendChild(target);
	const instance = mount(MediaSection, {
		target,
		props: { editor, mediaDataUrls: new Map([['media1.mp4', 'data:video/mp4;base64,AAAA']]) },
	});
	flushSync();
	cleanup = () => {
		unmount(instance);
		target.remove();
	};
	return target;
}

/** Simulate the `<video>` reporting a 20s duration, as `onloadedmetadata` does. */
function loadDuration(target: HTMLElement, seconds: number): void {
	const video = target.querySelector('video');
	if (!video) {
		throw new Error('expected a <video> element');
	}
	Object.defineProperty(video, 'duration', { value: seconds, configurable: true });
	video.dispatchEvent(new Event('loadedmetadata'));
	flushSync();
}

function trimEndInput(target: HTMLElement): HTMLInputElement {
	const label = Array.from(target.querySelectorAll('label')).find((node) =>
		(node.textContent ?? '').includes('Trim End'),
	);
	const input = label?.querySelector('input');
	if (!input) {
		throw new Error('missing "Trim End" field');
	}
	return input;
}

describe('mediaSection - shared playback options', () => {
	it('draws the shared playback fields and writes a media patch', () => {
		const editor = makeEditor(mediaElement());
		const target = mountSection(editor);
		expect(target.querySelector('[data-testid="media-playAcrossSlides"]')).toBeNull();
		const loop = target.querySelector<HTMLInputElement>('[data-testid="media-loop"]')!;
		loop.checked = true;
		loop.dispatchEvent(new Event('change', { bubbles: true }));
		flushSync();
		expect((editor.slides[0]!.elements[0] as MediaPptxElement).loop).toBeTruthy();
	});
});

describe('mediaSection - trim-end absolute conversion', () => {
	it('shows the end as 15 s for a 20 s clip with trimEndMs=5000', () => {
		const editor = makeEditor(mediaElement({ trimEndMs: 5000 }));
		const target = mountSection(editor);
		loadDuration(target, 20);

		expect(trimEndInput(target).value).toBe('15');
	});

	it('stores 5000 (p14:trim/@end) when the user types an absolute end of 15 s on a 20 s clip', () => {
		const editor = makeEditor(mediaElement({ trimEndMs: 0 }));
		const target = mountSection(editor);
		loadDuration(target, 20);

		const input = trimEndInput(target);
		input.value = '15';
		input.dispatchEvent(new Event('change', { bubbles: true }));
		flushSync();

		const patched = editor.slides[0]?.elements[0] as MediaPptxElement;
		expect(patched.trimEndMs).toBe(5000);
	});
});
