/**
 * external-content.service.test.ts: linked web content follows PowerPoint's default, per viewer.
 *
 * The deck model keeps its linked addresses; what this viewer may fetch is the shared
 * `resolveExternalContentView` over the host input, the Trust Center option and the consent for
 * the presentation now open. Enable Content redraws without reloading, and a collaborator's
 * slides count as much as a loaded deck. Same bare-injector harness as
 * `load-content-options-tracking.test.ts`.
 */
import {
	DestroyRef,
	Injector,
	signal,
	ɵChangeDetectionScheduler as ChangeDetectionScheduler,
	ɵEffectScheduler as EffectScheduler,
} from '@angular/core';
import JSZip from 'jszip';
import { linkedContentView } from 'ooxml-ui/pptx';
import { PptxHandler } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { buildingBlocksContractDeck } from '../../../shared/test-utils/building-blocks-contract.mjs';
import { linkedContentDeck } from '../../../shared/test-utils/external-content-contract.mjs';
import { EditorStateService } from './editor-state.service';
import { ExternalContentService } from './external-content.service';
import { LoadContentService } from './load-content.service';
import { ViewerOptionsService } from './viewer-options.service';

const ORIGIN = 'https://linked-content.test';

function harness() {
	const injector = Injector.create({
		providers: [
			{ provide: DestroyRef, useValue: { onDestroy: () => () => {} } },
			{ provide: ChangeDetectionScheduler, useValue: { notify: () => {} } },
			{
				provide: EffectScheduler,
				useValue: { add: () => {}, schedule: () => {}, remove: () => {}, flush: () => {} },
			},
			ViewerOptionsService,
			LoadContentService,
			EditorStateService,
			ExternalContentService,
		],
	});
	const loader = injector.get(LoadContentService);
	loader.bindHostGoogleFonts(() => false);
	const editor = injector.get(EditorStateService);
	return {
		loader,
		editor,
		options: injector.get(ViewerOptionsService),
		external: injector.get(ExternalContentService),
		/** Open a presentation as the viewer does: load it, then seed the editor from it. */
		async open(bytes: Uint8Array): Promise<void> {
			await loader.load(bytes);
			editor.setSlides(loader.slides(), loader.sections());
		},
	};
}

const linkedDeck = () => linkedContentDeck(PptxHandler, JSZip);
const plainDeck = () => buildingBlocksContractDeck(PptxHandler);
const links = (value: unknown) => JSON.stringify(value).includes(ORIGIN);

afterEach(() => {
	localStorage.clear();
	vi.restoreAllMocks();
});

describe('linked web content in an Angular viewer', () => {
	it('opens with the link in the model, nothing drawn from it, and the bar', async () => {
		const { editor, external, open } = harness();
		await open(await linkedDeck());
		expect(links(editor.slides())).toBeTruthy();
		expect(external.view()).toStrictEqual({ allowed: false, barVisible: true });
		expect(links(linkedContentView(editor.slides(), external.allowed()))).toBeFalsy();
	});

	it('enable Content draws it without another load and leaves the deck as it was', async () => {
		const { editor, external, open } = harness();
		await open(await linkedDeck());
		const slides = editor.slides();
		const parse = vi.spyOn(PptxHandler.prototype, 'load');
		external.enable();
		expect(external.view()).toStrictEqual({ allowed: true, barVisible: false });
		expect(links(linkedContentView(editor.slides(), external.allowed()))).toBeTruthy();
		expect(parse).not.toHaveBeenCalled();
		expect(editor.slides()).toBe(slides);
	});

	it('close keeps it withheld, and the next presentation asks again', async () => {
		const { external, open } = harness();
		await open(await linkedDeck());
		external.close();
		expect(external.view()).toStrictEqual({ allowed: false, barVisible: false });
		await open(await linkedDeck());
		expect(external.view()).toStrictEqual({ allowed: false, barVisible: true });
		external.enable();
		await open(await linkedDeck());
		expect(external.view()).toStrictEqual({ allowed: false, barVisible: true });
	});

	it('follows a live host value: false withholds without a bar, true draws without asking', async () => {
		const { loader, external, open } = harness();
		const host = signal<boolean | undefined>(false);
		loader.bindHostAllowExternalContent(() => host());
		await open(await linkedDeck());
		expect(external.view()).toStrictEqual({ allowed: false, barVisible: false });
		external.enable();
		expect(external.allowed(), 'host false wins over Enable Content').toBeFalsy();
		host.set(true);
		expect(external.view()).toStrictEqual({ allowed: true, barVisible: false });
	});

	it('the Trust Center option, off by default, draws it without asking when on', async () => {
		const { options, external, open } = harness();
		expect(options.options().trust.allowExternalContent).toBeFalsy();
		await open(await linkedDeck());
		options.setValue('trust', 'allowExternalContent', true);
		expect(external.view()).toStrictEqual({ allowed: true, barVisible: false });
	});

	it("shows the bar for a collaborator's slides that link, and keeps consent when slides are adopted", async () => {
		const room = harness();
		await room.open(await linkedDeck());
		const peer = harness();
		await peer.open(await plainDeck());
		expect(peer.external.view().barVisible).toBeFalsy();
		// What adopting a room's slides does: the editor takes them, nothing loads.
		peer.editor.setSlides(room.editor.slides());
		expect(peer.external.view()).toStrictEqual({ allowed: false, barVisible: true });
		peer.external.enable();
		peer.editor.setSlides(room.editor.slides());
		expect(peer.external.allowed()).toBeTruthy();
		expect(room.external.allowed(), 'each viewer decides for itself').toBeFalsy();
	});

	it('shows no bar for a deck without linked content', async () => {
		const { external, open } = harness();
		await open(await plainDeck());
		expect(external.view()).toStrictEqual({ allowed: false, barVisible: false });
	});
});
