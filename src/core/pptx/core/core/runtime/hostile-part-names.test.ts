import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { renameFixtureParts } from '../../../__tests__/helpers/rename-package-parts';
import { PresentationBuilder } from '../../builders/sdk';
import { PptxHandler } from '../../PptxHandler';
import type { PptxData, PptxElement, PptxSlide } from '../../types';

const PROBE = '"><i data-injected="1">';
const SAFE_ID = /^[A-Za-z0-9\-._~/%]+$/u;
const TEMPLATE_SHAPE =
	'<p:sp><p:nvSpPr><p:cNvPr id="90" name="Logo"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>';
const LAYOUT_PLACEHOLDER =
	'<p:sp><p:nvSpPr><p:cNvPr id="91" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="7"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr></p:sp>';
const RENAMED_FOLDERS = ['ppt/slides/', 'ppt/slideLayouts/', 'ppt/slideMasters/'];

/** A saved deck whose slide, layout and master each hold a shape of their own. */
async function deckWithTemplateShapes(): Promise<Uint8Array> {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	data.slides = [createSlide('Blank').addText('Kept text').build()];
	const zip = await JSZip.loadAsync(await handler.save(data.slides));
	handler.dispose();
	for (const path of Object.keys(zip.files)) {
		if (/^ppt\/(slideLayouts|slideMasters)\/[^/]+\.xml$/u.test(path)) {
			const xml = await zip.file(path)!.async('string');
			const shapes = path.includes('slideLayouts')
				? `${TEMPLATE_SHAPE}${LAYOUT_PLACEHOLDER}`
				: TEMPLATE_SHAPE;
			zip.file(path, xml.replace('</p:spTree>', `${shapes}</p:spTree>`));
		}
	}
	return zip.generateAsync({ type: 'uint8array' });
}

/** Give every slide, layout and master part a name holding markup. */
async function withHostilePartNames(bytes: Uint8Array): Promise<Uint8Array> {
	const zip = await JSZip.loadAsync(bytes);
	const names = new Map<string, string>();
	for (const path of Object.keys(zip.files)) {
		const folder = RENAMED_FOLDERS.find((candidate) => path.startsWith(candidate));
		if (folder && path.endsWith('.xml') && !path.includes('_rels')) {
			names.set(path, `${folder}p${PROBE}${path.slice(folder.length)}`);
		}
	}
	return renameFixtureParts(bytes, names);
}

function collectIds(elements: readonly PptxElement[] | undefined, into: string[]): void {
	for (const element of elements ?? []) {
		into.push(element.id);
		if (element.type === 'group') collectIds(element.children, into);
	}
}

async function allElementIds(handler: PptxHandler, data: PptxData): Promise<string[]> {
	const ids: string[] = [];
	for (const slide of data.slides) {
		collectIds(slide.elements, ids);
		collectIds(await handler.getTemplateElementsForSlide(slide.id), ids);
	}
	for (const master of data.slideMasters ?? []) {
		collectIds(master.elements, ids);
		for (const layout of master.layouts ?? []) collectIds(layout.elements, ids);
	}
	return ids;
}

async function load(bytes: Uint8Array): Promise<{ handler: PptxHandler; data: PptxData }> {
	const handler = new PptxHandler();
	const data = await handler.load(bytes.slice().buffer);
	return { handler, data };
}

function slideText(slide: PptxSlide | undefined): string {
	return JSON.stringify(slide?.elements ?? []);
}

describe('hostile part names', () => {
	it('never reach the ids of slide, layout or master elements', async () => {
		const { handler, data } = await load(
			await withHostilePartNames(await deckWithTemplateShapes()),
		);
		const ids = await allElementIds(handler, data);
		handler.dispose();
		expect(ids.filter((id) => id.includes('%3C')).length).toBeGreaterThanOrEqual(3);
		expect(ids.filter((id) => !SAFE_ID.test(id))).toEqual([]);
	});

	it('never reach the ids of placeholders a layout switch creates', async () => {
		const { handler, data } = await load(
			await withHostilePartNames(await deckWithTemplateShapes()),
		);
		const ids: string[] = [];
		for (const layout of handler.getLayoutOptions()) {
			const slide = await handler.applyLayoutToSlide(0, layout.path, data.slides);
			collectIds(slide.elements, ids);
		}
		handler.dispose();
		const generated = ids.filter((id) => id.startsWith('ph-'));
		expect(generated.length).toBeGreaterThan(0);
		expect(generated.filter((id) => !SAFE_ID.test(id))).toEqual([]);
		expect(generated.some((id) => id.includes('%3C'))).toBe(true);
	});

	it('still load, and survive a save and reload', async () => {
		const { handler, data } = await load(
			await withHostilePartNames(await deckWithTemplateShapes()),
		);
		expect(slideText(data.slides[0])).toContain('Kept text');
		const saved = await handler.save(data.slides);
		handler.dispose();
		const reopened = await load(saved);
		reopened.handler.dispose();
		expect(slideText(reopened.data.slides[0])).toContain('Kept text');
	});

	it('leave the ids of ordinary decks unchanged', async () => {
		const { handler, data } = await load(await deckWithTemplateShapes());
		const ids = await allElementIds(handler, data);
		handler.dispose();
		expect(ids).toContain('ppt/slides/slide1.xml-shape-0');
		expect(ids.some((id) => id.startsWith('layout-layout-shape-slideLayout'))).toBe(true);
		expect(ids.some((id) => id.startsWith('master-master-shape-slideMaster'))).toBe(true);
		expect(ids.some((id) => id.startsWith('slide-layout-slideLayout'))).toBe(true);
		expect(ids.some((id) => id.startsWith('slide-master-slideMaster'))).toBe(true);
	});
});
