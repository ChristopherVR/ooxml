import { flushPromises, mount } from '@vue/test-utils';
import type { VueWrapper } from '@vue/test-utils';
import JSZip from 'jszip';
import { PptxHandler } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import PowerPointViewer from './PowerPointViewer.vue';

/**
 * PowerPoint's default for linked pictures and media: nothing is downloaded on
 * open, a security bar offers Enable Content for that presentation, and
 * enabling reloads it with the links. A new presentation asks again.
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

/** The `allowExternalImages` each `PptxHandler.load` call so far received. */
function allowedPerLoad(load: LoadSpy): unknown[] {
	return load.mock.calls.map(
		(call) => (call[1] as { allowExternalImages?: boolean })?.allowExternalImages,
	);
}

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

const bar = () => mounted!.find('pptx-ui-external-content-bar');

function request(id: 'action' | 'close'): void {
	bar().element.dispatchEvent(new CustomEvent('external-content-request', { detail: { id } }));
}

describe('powerPointViewer linked external content', () => {
	it('blocks linked content on open and shows the security bar', async () => {
		const load = await mountDeck();
		expect(allowedPerLoad(load)).toStrictEqual([false]);
		expect(bar().exists()).toBeTruthy();
		expect(mounted!.html()).not.toContain(LINKED_ORIGIN);
	});

	it('reloads the presentation with its linked content on Enable Content and hides the bar', async () => {
		const load = await mountDeck();
		request('action');
		await settle(load, 2);
		expect(allowedPerLoad(load)).toStrictEqual([false, true]);
		expect(bar().exists()).toBeFalsy();
		expect(mounted!.html()).toContain(LINKED_ORIGIN);
	});

	it('asks again for a new presentation', async () => {
		// Built before the spy: building a deck parses one too.
		const next = await linkedDeck();
		const load = await mountDeck();
		request('action');
		await settle(load, 2);
		await mounted!.setProps({ content: next });
		await settle(load, 3);
		expect(allowedPerLoad(load)).toStrictEqual([false, true, false]);
		expect(bar().exists()).toBeTruthy();
	});

	it('close hides the bar and keeps the content blocked', async () => {
		const load = await mountDeck();
		request('close');
		await flushPromises();
		expect(bar().exists()).toBeFalsy();
		expect(load).toHaveBeenCalledOnce();
	});

	it('shows no bar when the host blocks linked content', async () => {
		const load = await mountDeck({ allowExternalContent: false });
		expect(allowedPerLoad(load)).toStrictEqual([false]);
		expect(bar().exists()).toBeFalsy();
	});

	it('downloads without a bar when the host allows linked content', async () => {
		const load = await mountDeck({ allowExternalContent: true });
		expect(allowedPerLoad(load)).toStrictEqual([true]);
		expect(bar().exists()).toBeFalsy();
	});
});
