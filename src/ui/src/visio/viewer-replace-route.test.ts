import { afterEach, expect, it } from 'vitest';
import { mountViewer } from './binding';
import { demoDocument } from 'ooxml-core/visio/ui';
import { replaceFixture } from './__fixtures__/replace-viewer';
import type { FindBar } from './viewer-search';

afterEach(() => document.body.replaceChildren());
it('routes Home Replace and CtrlH into shared occurrence mode, and CtrlF restores ordinary Find', async () => {
	const host = document.createElement('div');
	document.body.append(host);
	const viewer = mountViewer(host, { document: demoDocument });
	const root = viewer.element.shadowRoot!,
		viewport = root.querySelector<HTMLElement>('.viewport')!;
	const button = root.querySelector<HTMLElement & { disabled: boolean }>('[command="replace"]')!;
	const bar = root.querySelector<FindBar>('office-ui-find-bar')!;
	expect(button.disabled).toBe(true);
	await viewer.load(await replaceFixture());
	expect(button.disabled).toBe(false);
	button.shadowRoot!.querySelector('button')!.click();
	await bar.updateComplete;
	expect(bar.open).toBe(true);
	expect(bar.replaceMode).toBe(true);
	const input = bar.shadowRoot!.querySelector<HTMLInputElement>('input')!;
	input.value = 'CAT';
	input.dispatchEvent(new Event('input'));
	expect(bar.status).toBe('1 occurrence');
	viewport.dispatchEvent(
		new KeyboardEvent('keydown', { key: 'f', ctrlKey: true, bubbles: true, composed: true }),
	);
	expect(bar.replaceMode).toBe(false);
	expect(viewer.controller.state.search.results).toHaveLength(3);
	viewport.dispatchEvent(
		new KeyboardEvent('keydown', { key: 'h', ctrlKey: true, bubbles: true, composed: true }),
	);
	expect(bar.replaceMode).toBe(true);
	expect(bar.status).toBe('1 occurrence');
	const selection = viewer.controller.state.selectedShapes,
		bytes = viewer.controller.exportVsdx().bytes;
	input.dispatchEvent(
		new KeyboardEvent('keydown', { key: 'a', ctrlKey: true, bubbles: true, composed: true }),
	);
	input.dispatchEvent(
		new KeyboardEvent('keydown', { key: 'b', ctrlKey: true, bubbles: true, composed: true }),
	);
	expect(viewer.controller.state.selectedShapes).toBe(selection);
	expect(viewer.controller.exportVsdx().bytes).toEqual(bytes);
	viewer.destroy();
});

it('explains replacement source availability while original Find remains usable for model-only drawings', async () => {
	const host = document.createElement('div');
	document.body.append(host);
	const viewer = mountViewer(host, { document: demoDocument });
	const root = viewer.element.shadowRoot!,
		viewport = root.querySelector<HTMLElement>('.viewport')!,
		bar = root.querySelector<FindBar>('office-ui-find-bar')!;
	viewport.dispatchEvent(
		new KeyboardEvent('keydown', { key: 'h', ctrlKey: true, bubbles: true, composed: true }),
	);
	expect(bar.replaceMode).toBe(true);
	expect(bar.status).toBe('Open a .vsdx file to replace text');
	expect(bar.hasAttribute('disabled')).toBe(true);
	viewport.dispatchEvent(
		new KeyboardEvent('keydown', { key: 'f', ctrlKey: true, bubbles: true, composed: true }),
	);
	expect(bar.replaceMode).toBe(false);
	expect(bar.hasAttribute('disabled')).toBe(false);
	viewer.destroy();
});
