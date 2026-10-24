import fs from 'node:fs';
import path from 'node:path';

import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../core/PptxHandler';
import type { PptxAction, PptxElement, PptxSlide } from '../../core/types';
import { elementActionToPptxAction } from '../../core/utils/element-actions';
import { requireFixture } from '../require-fixture';

// One slide; its rectangle links to an external URL through rId2.
const fixturePath = requireFixture(
	path.resolve(__dirname, '../fixtures/e2e/hyperlink-action.pptx'),
);
const LOADED_URL = 'https://example.com/hyperlink-fixture';
const CHANGED_URL = 'https://example.org/changed';

async function open(): Promise<{ handler: PptxHandler; slides: PptxSlide[] }> {
	const bytes = fs.readFileSync(fixturePath);
	const handler = new PptxHandler();
	const data = await handler.load(
		bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
	);
	return { handler, slides: data.slides };
}

const linked = (slide: PptxSlide): PptxElement => slide.elements.find((el) => el.actionClick)!;

function withAction(slide: PptxSlide, el: PptxElement, actionClick: PptxAction): PptxSlide {
	return {
		...slide,
		elements: slide.elements.map((e) => (e === el ? ({ ...e, actionClick } as PptxElement) : e)),
	};
}

/** A slide the editor adds: no part yet, so the save gives it one. */
const newSlide = (elements: PptxElement[] = []): PptxSlide =>
	({ id: 'slide-new', rId: '', slideNumber: 2, elements }) as unknown as PptxSlide;

/** Save, then the saved slide's `a:hlinkClick` target as its relationship names it, and the reopened action. */
async function saved(handler: PptxHandler, slides: PptxSlide[], slideIndex: number) {
	const bytes = await handler.save(slides);
	const zip = await JSZip.loadAsync(bytes);
	const part = `ppt/slides/slide${slideIndex + 1}.xml`;
	const xml = await zip.file(part)!.async('string');
	const rels = await zip.file(part.replace('slides/', 'slides/_rels/') + '.rels')!.async('string');
	const rId = xml.match(/<a:hlinkClick[^>]*r:id="([^"]*)"/)?.[1];
	const target = rId
		? rels
				.match(new RegExp(`<Relationship [^>]*Id="${rId}"[^>]*>`))?.[0]
				.match(/Target="([^"]*)"/)?.[1]
		: undefined;
	const reopened = await new PptxHandler().load(bytes);
	const action = reopened.slides[slideIndex]!.elements.find((el) => el.actionClick)?.actionClick;
	return { rId, target, action };
}

describe('an element hyperlink saves the target the model holds', () => {
	it('an address edited after load saves the new address', async () => {
		const { handler, slides } = await open();
		const el = linked(slides[0]!);
		// What the shared hyperlink dialog patch writes: the loaded action, with a new address.
		expect(el.actionClick).toMatchObject({ rId: 'rId2', url: LOADED_URL });
		const edited = withAction(slides[0]!, el, { ...el.actionClick, url: CHANGED_URL });
		const result = await saved(handler, [edited], 0);
		expect(result.target).toBe(CHANGED_URL);
		expect(result.action?.url).toBe(CHANGED_URL);
	});

	it('an unedited link keeps its relationship', async () => {
		const { handler, slides } = await open();
		const el = linked(slides[0]!);
		const touched = withAction(slides[0]!, el, { ...el.actionClick, tooltip: 'Open it' });
		const result = await saved(handler, [touched], 0);
		expect(result).toMatchObject({ rId: 'rId2', target: LOADED_URL });
		expect(result.action).toMatchObject({ url: LOADED_URL, tooltip: 'Open it' });
	});

	it('a slide jump made in the editor saves the slide it points at', async () => {
		const { handler, slides } = await open();
		const el = linked(slides[0]!);
		const jump = elementActionToPptxAction({ trigger: 'click', type: 'slide', slideIndex: 1 })!;
		const result = await saved(handler, [withAction(slides[0]!, el, jump), newSlide()], 0);
		expect(result.target).toBe('slide2.xml');
		expect(result.action).toMatchObject({
			action: 'ppaction://hlinksldjump',
			targetSlideIndex: 1,
		});
		// Reopened, it is a loaded jump: saving again keeps its relationship.
		const reopened = new PptxHandler();
		const data = await reopened.load(
			await handler.save([withAction(slides[0]!, el, jump), newSlide()]),
		);
		const again = await saved(reopened, data.slides, 0);
		expect(again).toMatchObject({ rId: result.rId, target: 'slide2.xml' });
	});

	it('a linked element moved to another slide saves its address there', async () => {
		const { handler, slides } = await open();
		const el = linked(slides[0]!);
		// Its rId2 names a relationship of slide 1, which the new slide does not have.
		const moved = { ...el, id: 'moved-copy' } as PptxElement;
		const result = await saved(handler, [slides[0]!, newSlide([moved])], 1);
		expect(result.target).toBe(LOADED_URL);
		expect(result.action?.url).toBe(LOADED_URL);
	});
});
