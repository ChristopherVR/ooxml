// @vitest-environment jsdom
import { afterEach, expect, it, vi } from 'vitest';
import { act, createElement, StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import {
	DocumentCanvas,
	PageNavigation,
	StatusBar,
	Toolbar,
	useViewerBuildingBlocks,
	type BuildingBlockOptions,
} from './building-blocks';
import { createDocument } from '../../core/src/index';

afterEach(() => {
	document.body.replaceChildren();
	vi.unstubAllGlobals();
});

it('renders every library control from the props, keeps nodes on updates and honours host callbacks', async () => {
	vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
	const host = document.createElement('div');
	document.body.append(host);
	const root = createRoot(host);
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p', runs: [{ text: 'React composed words' }] }];
	const opened = vi.fn();
	let blocks!: ReturnType<typeof useViewerBuildingBlocks>;
	function Host(props: { options: BuildingBlockOptions }) {
		blocks = useViewerBuildingBlocks(props.options);
		return createElement(
			'div',
			null,
			createElement(DocumentCanvas, blocks.canvasProps),
			blocks.toolbarProps && createElement(Toolbar, blocks.toolbarProps),
			blocks.navigationProps && createElement(PageNavigation, blocks.navigationProps),
			blocks.statusProps && createElement(StatusBar, blocks.statusProps),
		);
	}
	const render = (options: BuildingBlockOptions) =>
		act(async () => root.render(createElement(StrictMode, null, createElement(Host, { options }))));
	try {
		await render({ documentModel: model, canEdit: true });
		const label = (name: string) => host.querySelector<HTMLElement>(`[aria-label="${name}"]`);
		const bold = label('Bold') as HTMLButtonElement;
		expect(bold.disabled).toBe(false);
		for (const name of ['Font', 'Font size', 'Align left', 'Align center', 'Numbered list'])
			expect(label(name), name).not.toBeNull();
		expect(label('Find')).toBeNull();
		expect(host.textContent).toContain('Page 1 of 1');
		expect(host.textContent).toContain('3 words');
		await act(async () => bold.click());
		expect(bold.getAttribute('aria-pressed')).toBe('true');
		expect(label('Bold')).toBe(bold);
		const zoom = label('Zoom') as HTMLSelectElement;
		await act(async () => {
			zoom.value = '150';
			zoom.dispatchEvent(new Event('change', { bubbles: true }));
		});
		expect(blocks.statusProps?.zoom).toBe(150);
		await render({ documentModel: model, onOpenFindReplace: opened, canEdit: true });
		const find = label('Find') as HTMLButtonElement;
		await act(async () => find.click());
		expect(opened).toHaveBeenCalledOnce();
		await render({ documentModel: model, onOpenFindReplace: opened, canEdit: false });
		expect(label('Bold')).toBe(bold);
		expect(bold.disabled).toBe(true);
		expect((label('Find') as HTMLButtonElement).disabled).toBe(false);
	} finally {
		await act(async () => root.unmount());
	}
});
