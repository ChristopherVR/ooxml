import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import JSZip from 'jszip';
import { PptxHandler } from 'pptx-viewer-core';
import { flushSync, mount, tick, unmount } from 'svelte';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import ExternalContentBar from '../components/ExternalContentBar.svelte';
import type { ViewerStateBag } from './create-viewer-state-types';
import CreateViewerStateHarness from './CreateViewerStateHarness.svelte';

/**
 * Linked external content follows PowerPoint's default in the Svelte binding:
 * nothing linked is referenced on open, the security bar offers Enable Content
 * for that presentation, and the host `allowExternalContent` switch overrides it.
 */

const ORIGIN = 'https://linked-content.test';
const FIXTURE = resolve(process.cwd(), '../../../../e2e/pptx/fixtures/sample-deck.pptx');
const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const LINKED_FILL = `<p:sp><p:nvSpPr><p:cNvPr id="901" name="Linked fill"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:blipFill><a:blip r:link="rIdLinkedFill"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:spPr></p:sp>`;

/** The sample deck with a shape on its first slide filled from a picture at {@link ORIGIN}. */
async function linkedDeck(): Promise<Uint8Array> {
	const zip = await JSZip.loadAsync(readFileSync(FIXTURE));
	const slide = await zip.file('ppt/slides/slide1.xml')!.async('string');
	zip.file('ppt/slides/slide1.xml', slide.replace('</p:spTree>', `${LINKED_FILL}</p:spTree>`));
	const relsPath = 'ppt/slides/_rels/slide1.xml.rels';
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

let cleanup: (() => void) | undefined;
beforeEach(() => localStorage.clear());
afterEach(() => {
	cleanup?.();
	cleanup = undefined;
	vi.restoreAllMocks();
});

function renderHarness(allowExternalContent?: boolean): ViewerStateBag {
	let captured: ViewerStateBag | undefined;
	const instance = mount(CreateViewerStateHarness, {
		target: document.createElement('div'),
		props: {
			editable: true,
			...(allowExternalContent === undefined ? {} : { allowExternalContent }),
			onready: (state: ViewerStateBag) => {
				captured = state;
			},
		},
	});
	cleanup = () => unmount(instance);
	flushSync();
	return captured!;
}

/** Load as a new presentation and let the editor seed from it. */
async function open(state: ViewerStateBag, bytes: Uint8Array): Promise<void> {
	await state.loader.load(bytes);
	await tick();
}

const references = (state: ViewerStateBag) => JSON.stringify(state.editor.slides).includes(ORIGIN);

describe('linked external content (Svelte)', () => {
	it('references nothing linked on open and shows the security bar', async () => {
		const state = renderHarness();
		const load = vi.spyOn(PptxHandler.prototype, 'load');
		await open(state, await linkedDeck());
		expect(load.mock.calls[0]?.[1]).toHaveProperty('allowExternalImages', false);
		expect(state.loader.hasLinkedWebContent).toBeTruthy();
		expect(references(state)).toBeFalsy();
		expect(state.externalContent.barVisible).toBeTruthy();
	});

	it('enabling reloads with the linked content, keeps the slide and hides the bar', async () => {
		const state = renderHarness();
		await open(state, await linkedDeck());
		state.viewer.goTo(2);
		const load = vi.spyOn(PptxHandler.prototype, 'load');
		await state.externalContent.enable();
		expect(load.mock.calls[0]?.[1]).toHaveProperty('allowExternalImages', true);
		expect(references(state)).toBeTruthy();
		expect(state.externalContent.barVisible).toBeFalsy();
		expect(state.viewer.current).toBe(2);
		expect(state.editor.dirty).toBeFalsy();
	});

	it('enabling keeps unsaved edits, still unsaved', async () => {
		const state = renderHarness();
		await open(state, await linkedDeck());
		const count = state.editor.slides.length;
		state.editor.commitSlides(state.editor.slides.slice(0, -1));
		expect(state.editor.dirty).toBeTruthy();
		await state.externalContent.enable();
		expect(state.editor.slides).toHaveLength(count - 1);
		expect(references(state)).toBeTruthy();
		expect(state.editor.dirty).toBeTruthy();
	});

	it('asks again for the next presentation', async () => {
		const state = renderHarness();
		await open(state, await linkedDeck());
		await state.externalContent.enable();
		expect(references(state)).toBeTruthy();
		await open(state, await linkedDeck());
		expect(state.loader.externalContentEnabled).toBeFalsy();
		expect(references(state)).toBeFalsy();
		expect(state.externalContent.barVisible).toBeTruthy();
	});

	it('closing hides the bar and keeps the content blocked', async () => {
		const state = renderHarness();
		await open(state, await linkedDeck());
		state.externalContent.close();
		expect(state.externalContent.barVisible).toBeFalsy();
		expect(references(state)).toBeFalsy();
	});

	it('host false blocks without the bar; host true loads without it', async () => {
		const blocked = renderHarness(false);
		await open(blocked, await linkedDeck());
		expect(references(blocked)).toBeFalsy();
		expect(blocked.externalContent.barVisible).toBeFalsy();
		cleanup?.();

		const allowed = renderHarness(true);
		await open(allowed, await linkedDeck());
		expect(references(allowed)).toBeTruthy();
		expect(allowed.externalContent.barVisible).toBeFalsy();
	});

	it('shows no bar for a deck without linked content', async () => {
		const state = renderHarness();
		await open(state, new Uint8Array(readFileSync(FIXTURE)));
		expect(state.loader.hasLinkedWebContent).toBeFalsy();
		expect(state.externalContent.barVisible).toBeFalsy();
	});
});

describe('externalContentBar', () => {
	it('routes Enable Content and Close from the shared element', () => {
		const onenable = vi.fn();
		const onclose = vi.fn();
		const target = document.createElement('div');
		document.body.append(target);
		const instance = mount(ExternalContentBar, { target, props: { onenable, onclose } });
		cleanup = () => {
			unmount(instance);
			target.remove();
		};
		flushSync();
		const host = target.querySelector('pptx-ui-external-content-bar')!;
		expect(host.shadowRoot?.textContent).toContain('SECURITY WARNING');
		host
			.shadowRoot!.querySelector<HTMLButtonElement>('[data-testid="pptx-external-content-action"]')!
			.click();
		host
			.shadowRoot!.querySelector<HTMLButtonElement>('[data-testid="pptx-external-content-close"]')!
			.click();
		expect(onenable).toHaveBeenCalledOnce();
		expect(onclose).toHaveBeenCalledOnce();
	});
});
