import { flushSync } from 'svelte';
import { afterEach, describe, expect, it } from 'vitest';

import type { EditorController } from '../editor/editor-controller.svelte';
import { EditorState } from '../editor/editor-state.svelte';
import { PresentationLoader } from './presentation-loader.svelte';
import { useViewerEffects } from './viewer-effects.svelte';
import { ViewerState } from './viewer-state.svelte';

afterEach(() => document.head.replaceChildren());

describe('useViewerEffects embedded fonts', () => {
	it('registers the deck embedded fonts as @font-face like the other bindings', () => {
		const loader = new PresentationLoader();
		const cleanup = $effect.root(() => {
			useViewerEffects({
				getSource: () => undefined,
				getEditable: () => true,
				getInitialSlide: () => 0,
				getTranslator: () => ((key: string) => key) as never,
				loader,
				viewer: new ViewerState(),
				editor: new EditorState({ getCurrent: () => 0, getHandler: () => null as never }),
				controller: { closeInline: () => undefined } as unknown as EditorController,
				getOnload: () => undefined,
				getOnerror: () => undefined,
				getOnslidechange: () => undefined,
				getGoogleFonts: () => false,
			});
		});
		flushSync();
		loader.embeddedFonts = [
			{
				name: 'Zzz Embedded Face',
				dataUrl: 'data:font/ttf;base64,AAEAAAAA',
				format: 'truetype',
				bold: false,
				italic: false,
			} as never,
		];
		flushSync();
		expect(document.getElementById('pptx-svelte-embedded-fonts')?.textContent).toContain(
			'font-family: "Zzz Embedded Face"',
		);
		cleanup();
		flushSync();
		expect(document.getElementById('pptx-svelte-embedded-fonts')).toBeNull();
	});
});
