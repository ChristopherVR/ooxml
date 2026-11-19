// @vitest-environment happy-dom
/**
 * Linked web content follows PowerPoint's default in `PowerPointViewer`: a deck
 * that links a picture to another server opens without referencing it and shows
 * the security bar; Enable Content reloads that presentation with the link,
 * Close keeps it blocked, a new presentation asks again, and the host's `false`
 * blocks without a bar.
 */
import JSZip from 'jszip';
import { translationsEn } from 'ooxml-ui/pptx/i18n';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

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

const ORIGIN = 'https://linked-content.test';
const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const LINKED_FILL = `<p:sp><p:nvSpPr><p:cNvPr id="901" name="Linked fill"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:blipFill><a:blip r:link="rIdLinkedFill"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:spPr></p:sp>`;

/** One slide whose shape is filled with a picture linked to {@link ORIGIN}. */
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
			`<Relationship Id="rIdLinkedFill" Type="${IMAGE_REL}" Target="${ORIGIN}/fill.png" TargetMode="External"/></Relationships>`,
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
const linkedReferenced = () => container.innerHTML.includes(ORIGIN);

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

// A whole viewer loads (twice for Enable Content) in each test.
describe('linked web content in PowerPointViewer', { timeout: 60_000 }, () => {
	it('opens blocked with the security bar, and Enable Content reloads with the link', async () => {
		await open(await linkedDeck());
		expect(loads).toStrictEqual([false]);
		expect(bar()).not.toBeNull();
		expect(linkedReferenced()).toBeFalsy();

		await request('action');
		await settle(() => loads.length === 2 && bar() === null && linkedReferenced());
		expect(loads).toStrictEqual([false, true]);
		expect(bar()).toBeNull();
		expect(linkedReferenced()).toBeTruthy();
	});

	it('keeps edits saved before Enable Content, and the next save writes them', async () => {
		await open(await linkedDeck(), undefined, { canEdit: true });
		const [first] = handle!.getElements(0);
		await act(async () => {
			handle!.updateElement(first!.id, { x: first!.x + 40 });
		});
		const saved = await handle!.getSaveSnapshot();
		await act(async () => {
			expect(handle!.acknowledgeSave(saved.revision).status).toBe('clean');
		});
		expect(handle!.isDirty()).toBe(false);
		await request('action');
		await settle(() => loads.length === 2 && linkedReferenced());
		expect(linkedReferenced()).toBeTruthy();
		expect(handle!.getElements(0)[0]!.x).toBe(first!.x + 40);
		expect(handle!.isDirty()).toBe(false);
		const next = await handle!.getSaveSnapshot();
		const slideXml = async (bytes: Uint8Array) =>
			(await JSZip.loadAsync(bytes)).file('ppt/slides/slide1.xml')!.async('string');
		const offset = (xml: string) => xml.match(/<a:off [^>]*>/u)?.[0];
		expect(offset(await slideXml(next.content))).toBe(offset(await slideXml(saved.content)));
		expect(await slideXml(next.content)).toContain('r:link="rIdLinkedFill"');
	});

	it('keeps unsaved edits unsaved through Enable Content', async () => {
		await open(await linkedDeck(), undefined, { canEdit: true });
		const [first] = handle!.getElements(0);
		await act(async () => {
			handle!.updateElement(first!.id, { x: first!.x + 40 });
		});
		await request('action');
		await settle(() => loads.length === 2 && linkedReferenced());
		expect(handle!.getElements(0)[0]!.x).toBe(first!.x + 40);
		expect(handle!.isDirty()).toBe(true);
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
		expect(loads).toStrictEqual([false]);
		expect(linkedReferenced()).toBeFalsy();
	});

	it('asks again for the next presentation', async () => {
		const deck = await linkedDeck();
		await open(deck);
		await request('action');
		await settle(() => loads.length === 2 && bar() === null);
		await open(deck.slice());
		expect(loads).toStrictEqual([false, true, false]);
		expect(bar()).not.toBeNull();
		expect(linkedReferenced()).toBeFalsy();
	});

	it('shows no bar when the host blocks linked content', async () => {
		await open(await linkedDeck(), false);
		expect(loads).toStrictEqual([false]);
		expect(bar()).toBeNull();
		expect(linkedReferenced()).toBeFalsy();
	});

	it('loads the link without a bar when the host allows it', async () => {
		await open(await linkedDeck(), true);
		expect(loads).toStrictEqual([true]);
		expect(bar()).toBeNull();
	});
});
