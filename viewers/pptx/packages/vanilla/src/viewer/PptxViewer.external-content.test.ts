import JSZip from 'jszip';
import { PptxHandler } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import {
	drawsLinkedContent,
	LINKED_FILL_URL,
	modelHasLinkedFill,
} from '../../../shared/test-utils/external-content-contract.mjs';
import { createPptxViewer } from './PptxViewer';
import type { PptxViewerInstance, PptxViewerOptions } from './types';

/**
 * Linked pictures follow PowerPoint's default: opening a deck downloads none of them and a
 * security bar offers Enable Content for that presentation only.
 */
const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const LINKED_FILL = `<p:sp><p:nvSpPr><p:cNvPr id="901" name="Linked fill"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:blipFill><a:blip r:link="rIdLinkedFill"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:spPr></p:sp>`;

/** A one-slide deck whose shape fill is a picture linked to {@link LINKED_FILL_URL}. */
async function linkedDeck(): Promise<Uint8Array> {
	const { handler, data } = await PptxHandler.create({ initialSlideCount: 1 });
	data.slides[0].isDirty = true;
	const zip = await JSZip.loadAsync(await handler.save(data.slides));
	handler.dispose();
	const slidePath = 'ppt/slides/slide1.xml';
	const relsPath = 'ppt/slides/_rels/slide1.xml.rels';
	const slide = await zip.file(slidePath)!.async('string');
	zip.file(slidePath, slide.replace('</p:spTree>', `${LINKED_FILL}</p:spTree>`));
	const rels = await zip.file(relsPath)!.async('string');
	zip.file(
		relsPath,
		rels.replace(
			'</Relationships>',
			`<Relationship Id="rIdLinkedFill" Type="${IMAGE_REL}" Target="${LINKED_FILL_URL}" TargetMode="External"/></Relationships>`,
		),
	);
	return zip.generateAsync({ type: 'uint8array' });
}

let active: PptxViewerInstance[] = [];

afterEach(async () => {
	for (const viewer of active) {
		viewer.destroy();
	}
	active = [];
	document.body.replaceChildren();
	localStorage.clear();
	vi.restoreAllMocks();
	await new Promise<void>((resolve) => {
		setTimeout(resolve, 0);
	});
});

function mount(options: PptxViewerOptions = {}) {
	const load = vi.spyOn(PptxHandler.prototype, 'load');
	const container = document.createElement('div');
	document.body.appendChild(container);
	const viewer = createPptxViewer(container, { googleFonts: false, ...options });
	active.push(viewer);
	const bar = () => container.querySelector('pptx-ui-external-content-bar');
	const state = () =>
		(viewer as unknown as { store: { get(): { slides: unknown; dirty: boolean } } }).store.get();
	/** Whether anything drawn (stage, thumbnails) would make the browser contact the linked origin. */
	const drawsOrigin = () => drawsLinkedContent(container);
	/** Whether the open deck's model keeps its links. */
	const modelLinks = () => modelHasLinkedFill(state().slides);
	return { container, viewer, bar, load, drawsOrigin, modelLinks, state };
}

function request(bar: Element, id: 'action' | 'close'): void {
	bar.dispatchEvent(new CustomEvent('external-content-request', { detail: { id } }));
}

describe('vanilla linked external content', () => {
	it('draws nothing linked on open, keeps the links in the deck and shows the security bar', async () => {
		const { viewer, bar, drawsOrigin, modelLinks } = mount();
		await viewer.loadFile(await linkedDeck());
		expect(drawsOrigin()).toBe(false);
		expect(modelLinks()).toBe(true);
		expect(bar()).not.toBeNull();
	});

	it('draws the linked content on Enable Content without reloading or changing the deck', async () => {
		const { viewer, bar, load, drawsOrigin, state } = mount();
		await viewer.loadFile(await linkedDeck());
		const loads = load.mock.calls.length;
		const slides = state().slides;
		request(bar()!, 'action');
		expect(drawsOrigin()).toBe(true);
		expect(bar()).toBeNull();
		expect(load.mock.calls.length).toBe(loads);
		expect(state().slides).toBe(slides);
		expect(state().dirty).toBe(false);
	});

	it('keeps unsaved edits and their undo through Enable Content', async () => {
		const { viewer, bar, drawsOrigin, state } = mount({ editable: true });
		await viewer.loadFile(await linkedDeck());
		const [first] = viewer.getElements(0);
		viewer.updateElement(first!.id, { x: first!.x + 40 } as never);
		request(bar()!, 'action');
		expect(drawsOrigin()).toBe(true);
		expect(viewer.getElements(0)[0]!.x).toBe(first!.x + 40);
		expect(state().dirty).toBe(true);
		expect(viewer.canUndo()).toBe(true);
		viewer.undo();
		expect(viewer.getElements(0)[0]!.x).toBe(first!.x);
		// Undo never takes the links away: they are the deck's content.
		expect(drawsOrigin()).toBe(true);
	});

	it('saves the pictures still linked after an edit made while they are blocked', async () => {
		const { viewer } = mount({ editable: true });
		await viewer.loadFile(await linkedDeck());
		const [first] = viewer.getElements(0);
		viewer.updateElement(first!.id, { x: first!.x + 40 } as never);
		const saved = await viewer.getSaveSnapshot();
		const zip = await JSZip.loadAsync(saved.content);
		expect(await zip.file('ppt/slides/slide1.xml')!.async('string')).toContain(
			'r:link="rIdLinkedFill"',
		);
		expect(await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string')).toContain(
			`Target="${LINKED_FILL_URL}" TargetMode="External"`,
		);
	});

	it('shows the bar without the toolbar', async () => {
		const { viewer, bar } = mount({ showToolbar: false });
		await viewer.loadFile(await linkedDeck());
		expect(bar()).not.toBeNull();
	});

	it('asks again for the next presentation', async () => {
		const { viewer, bar, drawsOrigin } = mount();
		await viewer.loadFile(await linkedDeck());
		request(bar()!, 'action');
		expect(drawsOrigin()).toBe(true);
		await viewer.loadFile(await linkedDeck());
		expect(drawsOrigin()).toBe(false);
		expect(bar()).not.toBeNull();
	});

	it('keeps the content blocked when the bar is closed', async () => {
		const { viewer, bar, drawsOrigin } = mount();
		await viewer.loadFile(await linkedDeck());
		request(bar()!, 'close');
		expect(bar()).toBeNull();
		expect(drawsOrigin()).toBe(false);
	});

	it('shows no bar when the host blocks linked content, and none it allows', async () => {
		const blocked = mount({ allowExternalContent: false });
		await blocked.viewer.loadFile(await linkedDeck());
		expect(blocked.drawsOrigin()).toBe(false);
		expect(blocked.bar()).toBeNull();

		vi.restoreAllMocks();
		const allowedByHost = mount({ allowExternalContent: true });
		await allowedByHost.viewer.loadFile(await linkedDeck());
		expect(allowedByHost.drawsOrigin()).toBe(true);
		expect(allowedByHost.bar()).toBeNull();
	});

	it('shows no bar for a deck that links nothing', async () => {
		const { handler, data } = await PptxHandler.create({ initialSlideCount: 1 });
		const bytes = await handler.save(data.slides);
		handler.dispose();
		const { viewer, bar } = mount();
		await viewer.loadFile(bytes);
		expect(bar()).toBeNull();
	});
});
