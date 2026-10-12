import { afterEach, expect, it, vi } from 'vitest';
import { parseVsdx } from 'ooxml-core/visio';
import { setupFormattingViewer } from './__fixtures__/formatting-viewer';
import { copySnapshotScene, demoDocument } from 'ooxml-core/visio/ui';
import { renderPage } from './render-svg';

afterEach(() => {
	document.body.replaceChildren();
	vi.restoreAllMocks();
});

it('retains safe accessibility metadata in interactive and static SVG snapshots', () => {
	const source = structuredClone(demoDocument);
	source.pages[0]!.accessibilityTitle = 'Accessible page';
	source.pages[0]!.alternativeText = 'Page description';
	const shape = source.pages[0]!.shapes[0]!;
	shape.accessibilityTitle = 'Status <script>';
	shape.alternativeText = 'First line\nSecond <img onerror="bad">';
	const snapshot = copySnapshotScene(source);
	shape.alternativeText = 'Changed after capture';
	for (const options of [{}, { static: true }]) {
		const rendered = renderPage(snapshot, snapshot.pages[0]!, options);
		expect(rendered.svg.getAttribute('aria-label')).toBe('Accessible page');
		expect(rendered.svg.querySelector(':scope > desc')?.textContent).toBe('Page description');
		expect(rendered.svg.textContent).toContain('First line\nSecond <img onerror="bad">');
		expect(rendered.svg.querySelector('script, img')).toBeNull();
		rendered.dispose();
	}
});

it('edits page alternatives without selecting a shape and restores them through history', async () => {
	const view = await setupFormattingViewer();
	expect(view.button('page-alt-text').disabled).toBe(false);
	view.press('page-alt-text');
	const dialog = view.root.querySelector<HTMLElement & { open: boolean }>('.alt-text-dialog')!;
	expect(dialog.getAttribute('heading')).toBe('Page Alt Text');
	view.root.querySelector<HTMLInputElement>('[name="alt-title"]')!.value = 'Process overview';
	view.root.querySelector<HTMLTextAreaElement>('[name="alt-description"]')!.value =
		'Arrive\nApprove';
	dialog
		.querySelector('[label="OK"]')!
		.shadowRoot!.querySelector<HTMLButtonElement>('button')!
		.click();
	await vi.waitFor(() => expect(dialog.open).toBe(false));
	expect(view.controller.state.document!.pages[0]!.alternativeText).toBe('Arrive\nApprove');
	const saved = await parseVsdx(view.controller.exportVsdx().bytes);
	expect(saved.pages[0]!.accessibilityTitle).toBe('Process overview');
	await view.controller.undo();
	expect(view.controller.state.document!.pages[0]!.alternativeText).toBeUndefined();
	view.dispose();
});

it('edits multiline native alternative text from Review, with undo and save/reopen', async () => {
	const view = await setupFormattingViewer();
	expect(view.button('alt-text').disabled).toBe(true);
	view.selection();
	view.press('alt-text');
	const dialog = view.root.querySelector<HTMLElement & { open: boolean }>('.alt-text-dialog')!;
	expect(dialog.open).toBe(true);
	view.root.querySelector<HTMLInputElement>('[name="alt-title"]')!.value = 'Service status';
	view.root.querySelector<HTMLTextAreaElement>('[name="alt-description"]')!.value =
		'Green box\nAll services operational';
	dialog
		.querySelector('[label="OK"]')!
		.shadowRoot!.querySelector<HTMLButtonElement>('button')!
		.click();
	await vi.waitFor(() => expect(dialog.open).toBe(false));
	expect(view.shape().alternativeText).toBe('Green box\nAll services operational');
	const reopened = await parseVsdx(view.controller.exportVsdx().bytes);
	expect(reopened.pages[0]!.shapes[0]!.accessibilityTitle).toBe('Service status');
	await view.controller.undo();
	expect(view.shape().alternativeText).toBeUndefined();
	await view.controller.redo();
	expect(view.shape().accessibilityTitle).toBe('Service status');
	view.press('page-alt-text');
	expect(view.root.querySelector<HTMLInputElement>('[name="alt-title"]')!.disabled).toBe(false);
	expect(view.root.querySelector<HTMLTextAreaElement>('[name="alt-description"]')!.disabled).toBe(
		false,
	);
	view.dispose();
});

it('reports its limited scope and opens the affected shape for correction', async () => {
	const view = await setupFormattingViewer();
	await view.controller.applyEdits([
		{ type: 'replace-plain-text', pageId: '1', shapeId: '1', text: '' },
	]);
	view.press('check-accessibility');
	const report = view.root.querySelector<HTMLElement & { open: boolean }>('.accessibility-dialog')!;
	expect(report.open).toBe(true);
	expect(report.textContent).toContain('Review reading order and color contrast separately');
	expect(report.textContent).toContain('1 graphic needs alternative text');
	report.querySelector<HTMLButtonElement>('button')!.click();
	expect(report.open).toBe(false);
	expect(view.root.querySelector<HTMLElement & { open: boolean }>('.alt-text-dialog')!.open).toBe(
		true,
	);
	view.dispose();
});
