// @vitest-environment jsdom
import { afterEach, expect, it, vi } from 'vitest';
import { createApp, defineComponent, h, nextTick } from 'vue';
import { DocumentCanvas, Toolbar, useViewerBuildingBlocks } from './building-blocks';
import { stubSpeech } from '../../bindings/src/speech-stub';
import { createDocument } from '../../core/src/index';

afterEach(() => {
	document.body.replaceChildren();
	vi.unstubAllGlobals();
});

it('exposes Read Aloud state, a control, and stops on unmount', async () => {
	const speech = stubSpeech();
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p', runs: [{ text: 'Vue words' }] }];
	const host = document.createElement('div');
	document.body.append(host);
	let blocks!: ReturnType<typeof useViewerBuildingBlocks>;
	const app = createApp(
		defineComponent({
			setup() {
				blocks = useViewerBuildingBlocks({ documentModel: model });
				return () =>
					h('div', [
						h(DocumentCanvas, { ...blocks.canvasProps }),
						blocks.toolbarProps.value && h(Toolbar, { toolbarProps: blocks.toolbarProps.value }),
					]);
			},
		}),
	);
	app.mount(host);
	await nextTick();
	expect(blocks.toolbarProps.value).toMatchObject({ canReadAloud: true, readingAloud: false });
	host.querySelector<HTMLButtonElement>('button[aria-label="Read aloud"]')!.click();
	await nextTick();
	expect(speech.spoken).toEqual(['Vue words']);
	expect(blocks.toolbarProps.value?.readingAloud).toBe(true);
	app.unmount();
	expect(speechSynthesis.cancel).toHaveBeenCalled();
});
