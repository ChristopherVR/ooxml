// @vitest-environment jsdom
import {
	createComponent,
	createEffect,
	createMemo,
	createRoot,
	createSignal,
	untrack,
} from 'solid-js';
import { render } from 'solid-js/web';
import { afterEach, expect, it, vi } from 'vitest';
import {
	DocumentCanvas,
	PageNavigation,
	StatusBar,
	Toolbar,
	createViewerBuildingBlocks,
	type BuildingBlockOptions,
} from './building-blocks';
import { createDocument } from '../../core/src/index';

afterEach(() => document.body.replaceChildren());

it('composes a surface with every library control, keeps nodes on updates and cleans up', async () => {
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p', runs: [{ text: 'Solid composed words' }] }];
	const host = document.createElement('div');
	document.body.append(host);
	const [options, setOptions] = createSignal<BuildingBlockOptions>({
		documentModel: model,
		canEdit: true,
	});
	let blocks!: ReturnType<typeof createViewerBuildingBlocks>;
	let dispose = () => {};
	createRoot((stop) => {
		dispose = stop;
		render(() => {
			blocks = createViewerBuildingBlocks(options);
			const shell = document.createElement('div');
			shell.append(createComponent(DocumentCanvas, blocks.canvasProps) as Node);
			const controls = document.createElement('div');
			shell.append(controls);
			const present = createMemo(() => blocks.toolbarProps() !== null);
			createEffect(() => {
				if (!present()) return;
				untrack(() =>
					controls.append(
						Toolbar({
							get toolbarProps() {
								return blocks.toolbarProps()!;
							},
						}) as Node,
						PageNavigation({
							get navigationProps() {
								return blocks.navigationProps()!;
							},
						}) as Node,
						StatusBar({
							get statusProps() {
								return blocks.statusProps()!;
							},
						}) as Node,
					),
				);
			});
			return shell;
		}, host);
	});
	try {
		await vi.waitFor(() => expect(host.innerHTML).toContain('data-control="bold"'));
		expect(host.querySelectorAll('docx-editor')).toHaveLength(1);
		const engine = host.querySelector('docx-editor');
		const bold = host.querySelector<HTMLButtonElement>('[data-control="bold"]')!;
		expect(bold.getAttribute('aria-pressed')).toBe('false');
		expect(host.querySelector('[data-control="fontFamily"]')).not.toBeNull();
		expect(host.querySelector('[data-control="alignCenter"]')).not.toBeNull();
		expect(host.querySelector('[data-control="numberList"]')).not.toBeNull();
		bold.click();
		expect(blocks.toolbarProps()?.bold).toBe(true);
		expect(host.textContent).toContain('Page 1 of 1');
		expect(host.textContent).toContain('3 words');
		setOptions((current) => ({ ...current, canEdit: false }));
		expect(blocks.mode()).toBe('view');
		expect(host.querySelector('docx-editor')).toBe(engine);
		expect(host.querySelector('[data-control="bold"]')).toBe(bold);
		expect(bold.disabled).toBe(true);
		expect(blocks.toolbarProps()?.onItalic()).toBe(false);
	} finally {
		dispose();
	}
	expect(host.querySelector('docx-editor')).toBeNull();
});
