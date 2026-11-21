// @vitest-environment happy-dom
/**
 * Linked web content follows PowerPoint's default in `PowerPointViewer`: a deck
 * that links a picture to another server opens without drawing it and shows the
 * security bar; Enable Content draws it without reloading (the deck, its unsaved
 * edits and undo history stay), Close keeps it blocked, a new presentation asks
 * again, and the host's `false` blocks without a bar. The model always keeps the
 * link, so a save writes it back.
 */
import JSZip from 'jszip';
import { translationsEn } from 'ooxml-ui/pptx/i18n';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import {
	drawsLinkedContent,
	LINKED_FILL_URL,
} from '../../../shared/test-utils/external-content-contract.mjs';

// oxlint-disable-next-line prefer-ending-with-an-expect
vi.mock<typeof import('react-i18next')>(import('react-i18next'), () => ({
	useTranslation: () => ({
		t: (key: string) => translationsEn[key] ?? key,
		i18n: {
			language: 'en',
			languages: ['en'],
			options: { resources: { en: {} } },
			changeLanguage: () => Promise.resolve(),
		},
	}),
}));

const { PptxHandler } = await import('pptx-viewer-core');
const { PowerPointViewer } = await import('../index');

const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const LINKED_FILL = `<p:sp><p:nvSpPr><p:cNvPr id="901" name="Linked fill"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:blipFill><a:blip r:link="rIdLinkedFill"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:spPr></p:sp>`;

/** One slide whose shape is filled with a picture linked to {@link LINKED_FILL_URL}. */
async function linkedDeck(): Promise<Uint8Array> {
	const { handler, data } = await PptxHandler.create({ initialSlideCount: 1 });
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

let container: HTMLDivElement;
let root: Root;
/** `allowExternalImages` of each load the viewer ran. */
let loads: (boolean | undefined)[];

beforeEach(() => {
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	window.happyDOM?.setViewport({ width: 1600, height: 950 });
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
	loads = [];
	const load = PptxHandler.prototype.load;
	vi.spyOn(PptxHandler.prototype, 'load').mockImplementation(function (this: unknown, ...args) {
		// The viewer's loads; building a deck (`PptxHandler.create`) loads without options.
		if (args[1] && 'allowExternalImages' in args[1]) {
			loads.push(args[1].allowExternalImages);
		}
		return load.apply(this, args);
	});
});

afterEach(() => {
	act(() => {
		root.unmount();
	});
	container.remove();
	vi.restoreAllMocks();
	globalThis.IS_REACT_ACT_ENVIRONMENT = false;
});

async function settle(done: () => boolean): Promise<void> {
	for (let attempt = 0; attempt < 150 && !done(); attempt += 1) {
		await act(async () => {
			await new Promise((resolve) => {
				setTimeout(resolve, 20);
			});
		});
	}
}

const bar = () => container.querySelector('pptx-ui-external-content-bar');
/** Whether anything the viewer draws would make the browser contact the linked server. */
const linkedDrawn = () => drawsLinkedContent(container);

type Handle = import('../types').PowerPointViewerHandle;
let handle: Handle | null = null;

async function open(
	content: Uint8Array,
	allowExternalContent?: boolean,
	extra: { canEdit?: boolean; showToolbar?: boolean } = {},
): Promise<void> {
	const onLoad = vi.fn();
	await act(async () => {
		root.render(
			<PowerPointViewer
				ref={(value: Handle | null) => {
					handle = value;
				}}
				content={content}
				onLoad={onLoad}
				{...extra}
				{...(allowExternalContent === undefined ? {} : { allowExternalContent })}
			/>,
		);
	});
	await settle(() => onLoad.mock.calls.length > 0);
	expect(onLoad).toHaveBeenCalledWith(expect.objectContaining({ slideCount: 1 }));
}

async function request(id: 'action' | 'close'): Promise<void> {
	await act(async () => {
		bar()!.dispatchEvent(new CustomEvent('external-content-request', { detail: { id } }));
	});
}

// A whole viewer loads in each test.
describe('linked web content in PowerPointViewer', { timeout: 60_000 }, () => {
	it('opens blocked with the security bar, and Enable Content draws the link without a reload', async () => {
		await open(await linkedDeck());
		// The model keeps the link; only drawing is gated.
		expect(loads).toStrictEqual([true]);
		expect(bar()).not.toBeNull();
		expect(linkedDrawn()).toBe(false);
		const slides = handle!.getSlides();

		await request('action');
		await settle(() => bar() === null && linkedDrawn());
		expect(loads).toStrictEqual([true]);
		expect(bar()).toBeNull();
		expect(linkedDrawn()).toBe(true);
		expect(handle!.getSlides()).toStrictEqual(slides);
	});

	it('keeps unsaved edits, the dirty state and undo through Enable Content', async () => {
		await open(await linkedDeck(), undefined, { canEdit: true });
		const [first] = handle!.getElements(0);
		await act(async () => {
			handle!.updateElement(first!.id, { x: first!.x + 40 });
		});
		await request('action');
		await settle(() => linkedDrawn());
		expect(handle!.getElements(0)[0]!.x).toBe(first!.x + 40);
		expect(handle!.isDirty()).toBe(true);
		await act(async () => {
			handle!.undo();
		});
		expect(handle!.getElements(0)[0]!.x).toBe(first!.x);
		expect(linkedDrawn()).toBe(true);
	});

	it('saves the picture still linked after an edit made while it is blocked', async () => {
		await open(await linkedDeck(), undefined, { canEdit: true });
		const [first] = handle!.getElements(0);
		await act(async () => {
			handle!.updateElement(first!.id, { x: first!.x + 40 });
		});
		const saved = await handle!.getSaveSnapshot();
		const zip = await JSZip.loadAsync(saved.content);
		expect(await zip.file('ppt/slides/slide1.xml')!.async('string')).toContain(
			'r:link="rIdLinkedFill"',
		);
		expect(await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string')).toContain(
			`Target="${LINKED_FILL_URL}" TargetMode="External"`,
		);
	});

	it('shows the bar without the toolbar', async () => {
		await open(await linkedDeck(), undefined, { showToolbar: false });
		expect(bar()).not.toBeNull();
	});

	it('closing hides the bar and keeps the content blocked', async () => {
		await open(await linkedDeck());
		await request('close');
		expect(bar()).toBeNull();
		await settle(() => false);
		expect(linkedDrawn()).toBe(false);
	});

	it('asks again for the next presentation', async () => {
		const deck = await linkedDeck();
		await open(deck);
		await request('action');
		await settle(() => bar() === null && linkedDrawn());
		await open(deck.slice());
		expect(bar()).not.toBeNull();
		expect(linkedDrawn()).toBe(false);
	});

	it('shows no bar when the host blocks linked content', async () => {
		await open(await linkedDeck(), false);
		expect(bar()).toBeNull();
		expect(linkedDrawn()).toBe(false);
	});

	it('draws the link without a bar when the host allows it', async () => {
		await open(await linkedDeck(), true);
		expect(bar()).toBeNull();
		await settle(() => linkedDrawn());
		expect(linkedDrawn()).toBe(true);
	});
});
