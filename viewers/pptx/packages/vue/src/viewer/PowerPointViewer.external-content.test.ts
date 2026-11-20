import { flushPromises, mount } from '@vue/test-utils';
import type { VueWrapper } from '@vue/test-utils';
import JSZip from 'jszip';
import { PptxHandler } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import PowerPointViewer from './PowerPointViewer.vue';
import type { PowerPointViewerExpose } from './types';

/**
 * PowerPoint's default for linked pictures and media: nothing is downloaded on
 * open, a security bar offers Enable Content for that presentation, and
 * enabling draws the links without reloading. The deck keeps its links
 * throughout. A new presentation asks again.
 */
const LINKED_ORIGIN = 'https://linked-content.test';
const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const LINKED_FILL_SHAPE = `<p:sp><p:nvSpPr><p:cNvPr id="901" name="Linked fill"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:blipFill><a:blip r:link="rIdLinkedFill"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:spPr></p:sp>`;

/** One slide whose shape fill is a picture linked to {@link LINKED_ORIGIN}. */
async function linkedDeck(): Promise<Uint8Array> {
	const { handler, createSlide } = await PptxHandler.create({ initialSlideCount: 0 });
	let bytes: Uint8Array;
	try {
		bytes = await handler.save([createSlide('Blank').build()]);
	} finally {
		handler.dispose();
	}
	const zip = await JSZip.loadAsync(bytes);
	const slidePath = 'ppt/slides/slide1.xml';
	const relsPath = 'ppt/slides/_rels/slide1.xml.rels';
	const slideXml = await zip.file(slidePath)!.async('string');
	zip.file(slidePath, slideXml.replace('</p:spTree>', `${LINKED_FILL_SHAPE}</p:spTree>`));
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

let mounted: VueWrapper | undefined;
afterEach(() => {
	mounted?.unmount();
	mounted = undefined;
	localStorage.clear();
	vi.restoreAllMocks();
});

type LoadSpy = ReturnType<typeof vi.spyOn<PptxHandler, 'load'>>;

async function settle(load: LoadSpy, loads: number): Promise<void> {
	await vi.waitFor(async () => {
		await flushPromises();
		expect(load).toHaveBeenCalledTimes(loads);
		expect(mounted?.find('.pptx-vue-loading').exists()).toBeFalsy();
	}, 15_000);
	await flushPromises();
}

async function mountDeck(props: Record<string, unknown> = {}): Promise<LoadSpy> {
	const content = await linkedDeck();
	const load = vi.spyOn(PptxHandler.prototype, 'load');
	// No Google Fonts request from the test environment once a deck renders.
	mounted = mount(PowerPointViewer, {
		props: { content, googleFonts: false, ...props },
		attachTo: document.body,
	});
	await settle(load, 1);
	return load;
}

const viewer = () => mounted!.vm as unknown as PowerPointViewerExpose;
const bar = () => mounted!.find('pptx-ui-external-content-bar');
/** Whether anything drawn (canvas, thumbnails) would make the browser contact the linked origin. */
const drawsOrigin = () => mounted!.html().includes(LINKED_ORIGIN);
/** Whether the open deck's model keeps its links. */
const modelLinks = () => JSON.stringify(viewer().getSlides()).includes(LINKED_ORIGIN);

async function request(id: 'action' | 'close'): Promise<void> {
	bar().element.dispatchEvent(new CustomEvent('external-content-request', { detail: { id } }));
	await flushPromises();
}

async function moveFirstElement(): Promise<{ id: string; x: number }> {
	const [first] = viewer().getElements(0);
	viewer().updateElement(first!.id, { x: first!.x + 40 });
	await flushPromises();
	return { id: first!.id, x: first!.x };
}

describe('powerPointViewer linked external content', () => {
	it('draws nothing linked on open, keeps the links in the deck and shows the security bar', async () => {
		const load = await mountDeck();
		// The model always keeps the links; drawing decides what is fetched.
		expect(load.mock.calls[0]?.[1]).toMatchObject({ allowExternalImages: true });
		expect(bar().exists()).toBeTruthy();
		expect(drawsOrigin()).toBeFalsy();
		expect(modelLinks()).toBeTruthy();
	});

	it('draws the linked content on Enable Content without reloading or changing the deck', async () => {
		const load = await mountDeck();
		const slides = JSON.stringify(viewer().getSlides());
		await request('action');
		expect(drawsOrigin()).toBeTruthy();
		expect(bar().exists()).toBeFalsy();
		expect(load).toHaveBeenCalledOnce();
		expect(JSON.stringify(viewer().getSlides())).toBe(slides);
		expect(viewer().isDirty()).toBeFalsy();
	});

	it('keeps unsaved edits and their undo through Enable Content', async () => {
		const load = await mountDeck({ canEdit: true });
		const first = await moveFirstElement();
		await request('action');
		expect(drawsOrigin()).toBeTruthy();
		expect(load).toHaveBeenCalledOnce();
		expect(viewer().getElements(0)[0]!.x).toBe(first.x + 40);
		expect(viewer().isDirty()).toBeTruthy();
		expect(viewer().canUndo()).toBeTruthy();
		viewer().undo();
		await flushPromises();
		expect(viewer().getElements(0)[0]!.x).toBe(first.x);
		// Undo never takes the links away: they are the deck's content.
		expect(drawsOrigin()).toBeTruthy();
	});

	it('saves the picture still linked after an edit made while it is blocked', async () => {
		await mountDeck({ canEdit: true });
		await moveFirstElement();
		const saved = await viewer().getSaveSnapshot();
		const zip = await JSZip.loadAsync(saved.content);
		await expect(zip.file('ppt/slides/slide1.xml')!.async('string')).resolves.toContain(
			'r:link="rIdLinkedFill"',
		);
		await expect(zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string')).resolves.toContain(
			`Target="${LINKED_ORIGIN}/fill.png" TargetMode="External"`,
		);
	});

	it('shows the bar without the toolbar', async () => {
		await mountDeck({ showToolbar: false });
		expect(bar().exists()).toBeTruthy();
	});

	it('asks again for a new presentation', async () => {
		// Built before the spy: building a deck parses one too.
		const next = await linkedDeck();
		const load = await mountDeck();
		await request('action');
		expect(drawsOrigin()).toBeTruthy();
		await mounted!.setProps({ content: next });
		await settle(load, 2);
		expect(drawsOrigin()).toBeFalsy();
		expect(bar().exists()).toBeTruthy();
	});

	it('keeps the content blocked when the bar is closed', async () => {
		const load = await mountDeck();
		await request('close');
		expect(bar().exists()).toBeFalsy();
		expect(drawsOrigin()).toBeFalsy();
		expect(load).toHaveBeenCalledOnce();
	});

	it('shows no bar and draws nothing linked when the host blocks linked content', async () => {
		await mountDeck({ allowExternalContent: false });
		expect(bar().exists()).toBeFalsy();
		expect(drawsOrigin()).toBeFalsy();
		expect(modelLinks()).toBeTruthy();
	});

	it('draws the linked content without a bar when the host allows it', async () => {
		await mountDeck({ allowExternalContent: true });
		expect(bar().exists()).toBeFalsy();
		expect(drawsOrigin()).toBeTruthy();
	});

	it('applies a live host change at once, without reloading', async () => {
		const load = await mountDeck();
		await mounted!.setProps({ allowExternalContent: true });
		await flushPromises();
		expect(drawsOrigin()).toBeTruthy();
		expect(bar().exists()).toBeFalsy();
		await mounted!.setProps({ allowExternalContent: false });
		await flushPromises();
		expect(drawsOrigin()).toBeFalsy();
		expect(bar().exists()).toBeFalsy();
		expect(load).toHaveBeenCalledOnce();
	});
});
