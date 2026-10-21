// @vitest-environment jsdom
import { afterEach, expect, it } from 'vitest';
import { createApp, defineComponent, h, nextTick, shallowRef } from 'vue';
import {
	DocumentCanvas,
	PageNavigation,
	StatusBar,
	Toolbar,
	useViewerBuildingBlocks,
	type BuildingBlockOptions,
} from './building-blocks';
import { createDocument } from '../../core/src/index';

afterEach(() => document.body.replaceChildren());

function document3() {
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p', runs: [{ text: 'Vue composed words' }] }];
	return model;
}

it('composes a surface with every library control, reacts to permission and disposes the engine', async () => {
	const host = document.createElement('div');
	document.body.append(host);
	const options = shallowRef<BuildingBlockOptions>({ documentModel: document3() });
	let blocks!: ReturnType<typeof useViewerBuildingBlocks>;
	const app = createApp(
		defineComponent({
			setup() {
				blocks = useViewerBuildingBlocks(options);
				return () =>
					h('div', [
						h(DocumentCanvas, { ...blocks.canvasProps }),
						blocks.toolbarProps.value && h(Toolbar, { toolbarProps: blocks.toolbarProps.value }),
						blocks.navigationProps.value &&
							h(PageNavigation, { navigationProps: blocks.navigationProps.value }),
						blocks.statusProps.value && h(StatusBar, { statusProps: blocks.statusProps.value }),
					]);
			},
		}),
	);
	app.mount(host);
	try {
		await nextTick();
		expect(host.querySelectorAll('docx-editor')).toHaveLength(1);
		const engine = host.querySelector('docx-editor');
		expect(host.querySelector('[role="toolbar"]')).not.toBeNull();
		const bold = host.querySelector<HTMLButtonElement>('button[aria-label="Bold"]')!;
		expect(bold.disabled).toBe(false);
		expect(host.querySelector('select[aria-label="Font"]')).not.toBeNull();
		expect(host.querySelector('button[aria-label="Align center"]')).not.toBeNull();
		expect(host.querySelector('button[aria-label="Bulleted list"]')).not.toBeNull();
		bold.click();
		await nextTick();
		expect(blocks.toolbarProps.value?.bold).toBe(true);
		// The same node is patched, so a focused control keeps its focus.
		expect(host.querySelector('button[aria-label="Bold"]')).toBe(bold);
		expect(bold.getAttribute('aria-pressed')).toBe('true');
		expect(host.textContent).toContain('Page 1 of 1');
		expect(host.textContent).toContain('3 words');
		options.value = { ...options.value, canEdit: false };
		await nextTick();
		expect(blocks.mode.value).toBe('view');
		expect(host.querySelector('docx-editor')).toBe(engine);
		expect(host.querySelector<HTMLButtonElement>('button[aria-label="Bold"]')!.disabled).toBe(true);
		expect(blocks.toolbarProps.value?.onItalic()).toBe(false);
	} finally {
		app.unmount();
	}
	expect(host.querySelector('docx-editor')).toBeNull();
});
