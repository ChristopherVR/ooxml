import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import JSZip from 'jszip';
import { PptxHandler } from 'pptx-viewer-core';
import { flushSync, mount, tick, unmount } from 'svelte';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import {
	drawsLinkedContent,
	LINKED_FILL_URL,
	modelHasLinkedFill,
} from '../../../../shared/test-utils/external-content-contract.mjs';
import ExternalContentBar from '../components/ExternalContentBar.svelte';
import PowerPointViewer from '../PowerPointViewer.svelte';
import type { ViewerStateBag } from './create-viewer-state-types';
import CreateViewerStateHarness from './CreateViewerStateHarness.svelte';

/**
 * Linked external content follows PowerPoint's default in the Svelte binding:
 * the deck keeps its links, nothing drawn references them on open, and Enable
 * Content draws them for this viewer without reloading or touching the deck.
 */

const FIXTURE = resolve(process.cwd(), '../../../../e2e/pptx/fixtures/sample-deck.pptx');
const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const LINKED_FILL = `<p:sp><p:nvSpPr><p:cNvPr id="901" name="Linked fill"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:blipFill><a:blip r:link="rIdLinkedFill"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:spPr></p:sp>`;

/** The sample deck with a shape on its first slide filled from a picture at {@link LINKED_FILL_URL}. */
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
			`<Relationship Id="rIdLinkedFill" Type="${IMAGE_REL}" Target="${LINKED_FILL_URL}" TargetMode="External"/></Relationships>`,
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

function renderHarness(props: { allowExternalContent?: boolean; filePath?: string } = {}) {
	let captured: ViewerStateBag | undefined;
	const instance = mount(CreateViewerStateHarness, {
		target: document.createElement('div'),
		props: {
			editable: true,
			...props,
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

const modelLinks = (state: ViewerStateBag) => modelHasLinkedFill(state.editor.slides);

describe('linked external content (Svelte)', () => {
	it('keeps the link in the model, draws nothing from it and shows the bar', async () => {
		const state = renderHarness();
		const load = vi.spyOn(PptxHandler.prototype, 'load');
		await open(state, await linkedDeck());
		expect(load.mock.calls[0]?.[1]).toHaveProperty('allowExternalImages', true);
		expect(modelLinks(state)).toBeTruthy();
		expect(state.externalContent.allowed).toBeFalsy();
		expect(state.externalContent.barVisible).toBeTruthy();
	});

	it('enabling redraws without a load and leaves the deck, slide and dirty state alone', async () => {
		const state = renderHarness();
		await open(state, await linkedDeck());
		state.viewer.goTo(2);
		const slides = state.editor.slides;
		const json = JSON.stringify(slides);
		const load = vi.spyOn(PptxHandler.prototype, 'load');
		state.externalContent.enable();
		flushSync();
		expect(load).not.toHaveBeenCalled();
		expect(state.externalContent.allowed).toBeTruthy();
		expect(state.externalContent.barVisible).toBeFalsy();
		expect(state.editor.slides).toBe(slides);
		expect(JSON.stringify(state.editor.slides)).toBe(json);
		expect(state.viewer.current).toBe(2);
		expect(state.editor.dirty).toBeFalsy();
	});

	it('enabling keeps unsaved edits and their undo history', async () => {
		const state = renderHarness();
		await open(state, await linkedDeck());
		const count = state.editor.slides.length;
		state.editor.commitSlides(state.editor.slides.slice(0, -1));
		expect(state.editor.dirty).toBeTruthy();
		state.externalContent.enable();
		flushSync();
		expect(state.editor.slides).toHaveLength(count - 1);
		expect(state.editor.dirty).toBeTruthy();
		expect(state.editor.canUndo).toBeTruthy();
		state.editor.undo();
		flushSync();
		expect(state.editor.slides).toHaveLength(count);
		expect(state.externalContent.allowed).toBeTruthy();
	});

	it('autosave still schedules a snapshot for an edit made after Enable Content', async () => {
		const state = renderHarness({ filePath: 'linked.pptx' });
		await open(state, await linkedDeck());
		flushSync();
		state.externalContent.enable();
		flushSync();
		await tick();
		expect(state.autosaveCtl.isDirty).toBeFalsy();
		state.editor.commitSlides(state.editor.slides.slice(0, -1));
		flushSync();
		expect(state.autosaveCtl.isDirty).toBeTruthy();
	});

	it('a save after an edit made while blocked keeps the link', async () => {
		const state = renderHarness();
		await open(state, await linkedDeck());
		state.editor.commitSlides(state.editor.slides.slice(0, -1));
		const saved = await state.editingApi.getSaveSnapshot();
		const zip = await JSZip.loadAsync(saved.content);
		const rels = await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string');
		const slide = await zip.file('ppt/slides/slide1.xml')!.async('string');
		const relationship = rels.match(/<Relationship [^>]*Id="rIdLinkedFill"[^>]*>/u)?.[0];
		expect(relationship).toContain(`Target="${LINKED_FILL_URL}"`);
		expect(relationship).toContain('TargetMode="External"');
		expect(slide).toContain('r:link=');
	});

	it('closing hides the bar and keeps the content blocked', async () => {
		const state = renderHarness();
		await open(state, await linkedDeck());
		state.externalContent.close();
		flushSync();
		expect(state.externalContent.barVisible).toBeFalsy();
		expect(state.externalContent.allowed).toBeFalsy();
	});

	it('asks again for the next presentation', async () => {
		const state = renderHarness();
		await open(state, await linkedDeck());
		state.externalContent.enable();
		flushSync();
		expect(state.externalContent.allowed).toBeTruthy();
		await open(state, await linkedDeck());
		expect(state.externalContent.allowed).toBeFalsy();
		expect(state.externalContent.barVisible).toBeTruthy();
	});

	it('applies the Trust Center option at once', async () => {
		const state = renderHarness();
		await open(state, await linkedDeck());
		state.optionsState.setValue('trust', 'allowExternalContent', true);
		flushSync();
		expect(state.externalContent.allowed).toBeTruthy();
		expect(state.externalContent.barVisible).toBeFalsy();
	});

	it('host false blocks without the bar; host true draws without it', async () => {
		const blocked = renderHarness({ allowExternalContent: false });
		await open(blocked, await linkedDeck());
		expect(modelLinks(blocked)).toBeTruthy();
		expect(blocked.externalContent.allowed).toBeFalsy();
		expect(blocked.externalContent.barVisible).toBeFalsy();
		cleanup?.();

		const allowed = renderHarness({ allowExternalContent: true });
		await open(allowed, await linkedDeck());
		expect(allowed.externalContent.allowed).toBeTruthy();
		expect(allowed.externalContent.barVisible).toBeFalsy();
	});

	it('shows no bar for a deck without linked content', async () => {
		const state = renderHarness();
		await open(state, new Uint8Array(readFileSync(FIXTURE)));
		expect(state.externalContent.barVisible).toBeFalsy();
	});
});

describe('powerPointViewer linked content', () => {
	it('draws the linked picture only after Enable Content, without reloading', async () => {
		const source = await linkedDeck();
		const target = document.createElement('div');
		document.body.append(target);
		const onload = vi.fn();
		const instance = mount(PowerPointViewer, { target, props: { source, editable: true, onload } });
		cleanup = () => {
			unmount(instance);
			target.remove();
		};
		flushSync();
		await vi.waitFor(() => expect(onload).toHaveBeenCalledOnce(), { timeout: 15000 });
		flushSync();
		const bar = () => target.querySelector('pptx-ui-external-content-bar');
		await vi.waitFor(() => expect(bar()).not.toBeNull());
		expect(drawsLinkedContent(target)).toBe(false);

		const load = vi.spyOn(PptxHandler.prototype, 'load');
		bar()!.dispatchEvent(
			new CustomEvent('external-content-request', { detail: { id: 'action' }, bubbles: true }),
		);
		flushSync();
		await vi.waitFor(() => expect(drawsLinkedContent(target)).toBe(true));
		expect(bar()).toBeNull();
		expect(load).not.toHaveBeenCalled();
	}, 30000);
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
