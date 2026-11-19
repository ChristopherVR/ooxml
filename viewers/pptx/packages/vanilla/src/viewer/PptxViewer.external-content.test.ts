import JSZip from 'jszip';
import { PptxHandler } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { createPptxViewer } from './PptxViewer';
import type { PptxViewerInstance, PptxViewerOptions } from './types';

/**
 * Linked pictures follow PowerPoint's default: opening a deck downloads none of them and a
 * security bar offers Enable Content for that presentation only.
 */
const LINKED_ORIGIN = 'https://linked-content.test';
const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const LINKED_FILL = `<p:sp><p:nvSpPr><p:cNvPr id="901" name="Linked fill"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:blipFill><a:blip r:link="rIdLinkedFill"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:spPr></p:sp>`;

/** A one-slide deck whose shape fill is a picture linked to {@link LINKED_ORIGIN}. */
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
			`<Relationship Id="rIdLinkedFill" Type="${IMAGE_REL}" Target="${LINKED_ORIGIN}/fill.png" TargetMode="External"/></Relationships>`,
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
	/** Whether the last load could download linked content (core may call `load` more than once per open). */
	const allowed = () => load.mock.lastCall?.[1]?.allowExternalImages === true;
	const state = () =>
		(viewer as unknown as { store: { get(): { slides: unknown; dirty: boolean } } }).store.get();
	/** Whether any element of the open deck references the linked origin. */
	const referencesOrigin = () => JSON.stringify(state().slides).includes(LINKED_ORIGIN);
	return { container, viewer, bar, allowed, referencesOrigin, state };
}

function request(bar: Element, id: 'action' | 'close'): void {
	bar.dispatchEvent(new CustomEvent('external-content-request', { detail: { id } }));
}

describe('vanilla linked external content', () => {
	it('blocks a linked deck on open and shows the security bar', async () => {
		const { viewer, bar, allowed, referencesOrigin } = mount();
		await viewer.loadFile(await linkedDeck());
		expect(allowed()).toBeFalsy();
		expect(referencesOrigin()).toBeFalsy();
		expect(bar()).not.toBeNull();
	});

	it('reloads the presentation with its linked content on Enable Content', async () => {
		const { viewer, bar, allowed, referencesOrigin, state } = mount();
		await viewer.loadFile(await linkedDeck());
		request(bar()!, 'action');
		await vi.waitFor(() => expect(allowed()).toBeTruthy());
		await vi.waitFor(() => expect(referencesOrigin()).toBeTruthy());
		expect(bar()).toBeNull();
		expect(state().dirty).toBeFalsy();
	});

	it('asks again for the next presentation', async () => {
		const { viewer, bar, allowed } = mount();
		await viewer.loadFile(await linkedDeck());
		request(bar()!, 'action');
		await vi.waitFor(() => expect(allowed()).toBeTruthy());
		await viewer.loadFile(await linkedDeck());
		expect(allowed()).toBeFalsy();
		expect(bar()).not.toBeNull();
	});

	it('keeps the content blocked when the bar is closed', async () => {
		const { viewer, bar, allowed, referencesOrigin } = mount();
		await viewer.loadFile(await linkedDeck());
		request(bar()!, 'close');
		expect(bar()).toBeNull();
		expect(allowed()).toBeFalsy();
		expect(referencesOrigin()).toBeFalsy();
	});

	it('shows no bar when the host blocks linked content, and none it allows', async () => {
		const blocked = mount({ allowExternalContent: false });
		await blocked.viewer.loadFile(await linkedDeck());
		expect(blocked.allowed()).toBeFalsy();
		expect(blocked.bar()).toBeNull();

		vi.restoreAllMocks();
		const allowedByHost = mount({ allowExternalContent: true });
		await allowedByHost.viewer.loadFile(await linkedDeck());
		expect(allowedByHost.allowed()).toBeTruthy();
		expect(allowedByHost.referencesOrigin()).toBeTruthy();
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
