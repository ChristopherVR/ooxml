/**
 * external-content-consent.test.ts: linked pictures follow PowerPoint's default.
 *
 * A deck that links a picture from the web loads with the link left out and
 * reports it, so the viewer can show the security bar; Enable Content reloads
 * that presentation with its linked content; the next presentation asks again.
 * Same bare-injector harness as `load-content-options-tracking.test.ts`.
 */
import {
	DestroyRef,
	Injector,
	ɵChangeDetectionScheduler as ChangeDetectionScheduler,
	ɵEffectScheduler as EffectScheduler,
} from '@angular/core';
import JSZip from 'jszip';
import { PptxHandler } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { LoadContentService } from './load-content.service';
import { ViewerOptionsService } from './viewer-options.service';

const LINKED_URL = 'https://external-content.test/fill.png';
const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const LINKED_FILL_SHAPE = `<p:sp><p:nvSpPr><p:cNvPr id="901" name="Linked fill"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:blipFill><a:blip r:link="rIdLinkedFill"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:spPr></p:sp>`;

function harness(): { loader: LoadContentService; options: ViewerOptionsService } {
	const injector = Injector.create({
		providers: [
			{ provide: DestroyRef, useValue: { onDestroy: () => () => {} } },
			{ provide: ChangeDetectionScheduler, useValue: { notify: () => {} } },
			{
				provide: EffectScheduler,
				useValue: { add: () => {}, schedule: () => {}, remove: () => {}, flush: () => {} },
			},
			{ provide: ViewerOptionsService, useClass: ViewerOptionsService },
			{ provide: LoadContentService, useClass: LoadContentService },
		],
	});
	return { loader: injector.get(LoadContentService), options: injector.get(ViewerOptionsService) };
}

async function deckBytes(linked: boolean): Promise<Uint8Array> {
	const { handler, data } = await PptxHandler.create({ title: 'Deck', initialSlideCount: 1 });
	const zip = await JSZip.loadAsync(await handler.save(data.slides));
	handler.dispose();
	if (linked) {
		const slidePath = 'ppt/slides/slide1.xml';
		const relsPath = 'ppt/slides/_rels/slide1.xml.rels';
		const slideXml = await zip.file(slidePath)!.async('string');
		zip.file(slidePath, slideXml.replace('</p:spTree>', `${LINKED_FILL_SHAPE}</p:spTree>`));
		const rels = await zip.file(relsPath)!.async('string');
		zip.file(
			relsPath,
			rels.replace(
				'</Relationships>',
				`<Relationship Id="rIdLinkedFill" Type="${IMAGE_REL}" Target="${LINKED_URL}" TargetMode="External"/></Relationships>`,
			),
		);
	}
	return zip.generateAsync({ type: 'uint8array' });
}

const referencesLink = (loader: LoadContentService) =>
	JSON.stringify(loader.slides()).includes(LINKED_URL);

afterEach(() => {
	localStorage.clear();
	vi.restoreAllMocks();
});

describe('linked external content in the Angular loader', () => {
	it('loads a linked deck without its link and reports it for the security bar', async () => {
		const { loader } = harness();
		const parse = vi.spyOn(PptxHandler.prototype, 'load');
		await loader.load(await deckBytes(true));
		expect(parse.mock.calls.at(-1)?.[1]).toMatchObject({ allowExternalImages: false });
		expect(loader.hasLinkedWebContent()).toBeTruthy();
		expect(loader.externalContentAllowed()).toBeFalsy();
		expect(referencesLink(loader)).toBeFalsy();
	});

	it('enable Content reloads that presentation with its linked content', async () => {
		const { loader } = harness();
		const bytes = await deckBytes(true);
		await loader.load(bytes);
		const parse = vi.spyOn(PptxHandler.prototype, 'load');
		await loader.enableExternalContent(bytes);
		expect(parse.mock.calls.at(-1)?.[1]).toMatchObject({ allowExternalImages: true });
		expect(loader.externalContent()).toStrictEqual({ enabled: true, closed: false });
		expect(loader.externalContentAllowed()).toBeTruthy();
		expect(loader.lastLoadWasExternalContentReload).toBeTruthy();
		expect(referencesLink(loader)).toBeTruthy();
	});

	it('asks again for the next presentation', async () => {
		const { loader } = harness();
		const bytes = await deckBytes(true);
		await loader.load(bytes);
		await loader.enableExternalContent(bytes);
		await loader.load(await deckBytes(true));
		expect(loader.externalContent()).toStrictEqual({ enabled: false, closed: false });
		expect(loader.externalContentAllowed()).toBeFalsy();
		expect(loader.lastLoadWasExternalContentReload).toBeFalsy();
		expect(referencesLink(loader)).toBeFalsy();
	});

	it('close keeps the content blocked for this presentation only', async () => {
		const { loader } = harness();
		await loader.load(await deckBytes(true));
		loader.closeExternalContentBar();
		expect(loader.externalContent()).toStrictEqual({ enabled: false, closed: true });
		await loader.load(await deckBytes(true));
		expect(loader.externalContent().closed).toBeFalsy();
	});

	it('a host false wins over Enable Content', async () => {
		const { loader } = harness();
		loader.bindHostAllowExternalContent(() => false);
		const bytes = await deckBytes(true);
		await loader.load(bytes);
		await loader.enableExternalContent(bytes);
		expect(loader.externalContentAllowed()).toBeFalsy();
		expect(referencesLink(loader)).toBeFalsy();
	});

	it('the Trust Center option, off by default, allows it without asking when on', async () => {
		const { loader, options } = harness();
		expect(options.options().trust.allowExternalContent).toBeFalsy();
		options.setValue('trust', 'allowExternalContent', true);
		await loader.load(await deckBytes(true));
		expect(loader.externalContentAllowed()).toBeTruthy();
		expect(referencesLink(loader)).toBeTruthy();
	});

	it('reports nothing for a deck without linked content', async () => {
		const { loader } = harness();
		await loader.load(await deckBytes(false));
		expect(loader.hasLinkedWebContent()).toBeFalsy();
	});
});
