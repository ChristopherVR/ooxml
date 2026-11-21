/**
 * PowerPoint's security bar in the Angular viewer: a deck that links a picture to the web opens
 * with the link in its model and nothing drawn from it. Enable Content draws it without loading
 * the deck again, so the slides, unsaved edits, dirty state and undo history carry on, and a save
 * keeps the link. Close keeps it withheld, the next presentation asks again, and the host
 * `allowExternalContent` input decides outright when set.
 */
import { TestBed } from '@angular/core/testing';
import type { ComponentFixture } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { provideTranslateService } from '@ngx-translate/core';
import JSZip from 'jszip';
import { PptxHandler } from 'pptx-viewer-core';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';

import {
	drawsLinkedContent,
	linkedContentDeck,
	LINKED_FILL_URL,
	modelHasLinkedFill,
} from '../../../shared/test-utils/external-content-contract.mjs';
import { resolveViewerComponentResources } from './component-resources.test-support';
import { PowerPointViewerComponent } from './power-point-viewer.component';
import { declareSignalFields } from './signal-inputs.test-support';

beforeAll(async () => {
	declareSignalFields(PowerPointViewerComponent, import.meta.dirname);
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	await resolveViewerComponentResources();
});

afterEach(() => {
	TestBed.resetTestingModule();
	document.body.replaceChildren();
	localStorage.clear();
	vi.restoreAllMocks();
});

type Viewer = ComponentFixture<PowerPointViewerComponent>;

async function settle(fixture: Viewer, done: () => boolean = () => true): Promise<void> {
	for (let attempt = 0; attempt < 200; attempt += 1) {
		fixture.detectChanges();
		await fixture.whenStable();
		await new Promise((resolve) => {
			setTimeout(resolve, 20);
		});
		if (done()) {
			break;
		}
	}
	fixture.detectChanges();
}

const root = (fixture: Viewer) => fixture.nativeElement as HTMLElement;
const bar = (fixture: Viewer) => root(fixture).querySelector('pptx-ui-external-content-bar');
/** Whether anything the viewer draws would contact the linked server. */
const drawn = (fixture: Viewer) => drawsLinkedContent(root(fixture));
const modelLinks = (fixture: Viewer) => modelHasLinkedFill(fixture.componentInstance.getSlides());
const shown = (fixture: Viewer) => root(fixture).querySelector('[data-pptx-element]') !== null;

async function open(inputs: Record<string, unknown> = {}): Promise<Viewer> {
	TestBed.configureTestingModule({ providers: [provideTranslateService()] });
	const fixture = TestBed.createComponent(PowerPointViewerComponent);
	document.body.append(root(fixture));
	fixture.componentRef.setInput('googleFonts', false);
	for (const [name, value] of Object.entries(inputs)) {
		fixture.componentRef.setInput(name, value);
	}
	fixture.componentRef.setInput('content', await linkedContentDeck(PptxHandler, JSZip));
	await settle(fixture, () => shown(fixture));
	return fixture;
}

function request(fixture: Viewer, id: 'action' | 'close'): void {
	bar(fixture)!.dispatchEvent(new CustomEvent('external-content-request', { detail: { id } }));
}

async function part(bytes: Uint8Array, path: string): Promise<string> {
	return (await JSZip.loadAsync(bytes)).file(path)!.async('string');
}

describe('angular viewer linked web content', { timeout: 60_000 }, () => {
	it('opens with the link in the model, nothing drawn from it, and the bar', async () => {
		const fixture = await open();
		expect(modelLinks(fixture)).toBeTruthy();
		expect(drawn(fixture)).toBeFalsy();
		expect(bar(fixture)).not.toBeNull();
	});

	it('enable Content draws it without a reload, keeping edits, dirty state and undo', async () => {
		const fixture = await open({ canEditInput: true });
		const viewer = fixture.componentInstance;
		const [first] = viewer.getElements(0);
		viewer.updateElement(first!.id, { x: first!.x + 40 });
		await settle(fixture);
		const slides = viewer.getSlides();
		const parse = vi.spyOn(PptxHandler.prototype, 'load');
		request(fixture, 'action');
		await settle(fixture, () => drawn(fixture));
		expect(drawn(fixture)).toBeTruthy();
		expect(bar(fixture)).toBeNull();
		expect(parse).not.toHaveBeenCalled();
		expect(JSON.stringify(viewer.getSlides())).toBe(JSON.stringify(slides));
		expect(viewer.getElements(0)[0]!.x).toBe(first!.x + 40);
		expect(viewer.isDirty()).toBeTruthy();
		expect(viewer.canUndo()).toBeTruthy();
		viewer.undo();
		await settle(fixture);
		expect(viewer.getElements(0)[0]!.x).toBe(first!.x);
		expect(drawn(fixture), 'undo after enabling keeps the content drawn').toBeTruthy();
	});

	it('saves an edit made while blocked with the link and its external relationship', async () => {
		const fixture = await open({ canEditInput: true });
		const viewer = fixture.componentInstance;
		const [first] = viewer.getElements(0);
		viewer.updateElement(first!.id, { x: first!.x + 40 });
		await settle(fixture);
		const { content } = await viewer.getSaveSnapshot();
		await expect(part(content, 'ppt/slides/slide1.xml')).resolves.toContain(
			'r:link="rIdLinkedFill"',
		);
		const relationship = (await part(content, 'ppt/slides/_rels/slide1.xml.rels')).match(
			/<Relationship [^>]*Id="rIdLinkedFill"[^>]*>/u,
		)?.[0];
		expect(relationship).toContain(`Target="${LINKED_FILL_URL}"`);
		expect(relationship).toContain('TargetMode="External"');
	});

	it('close keeps it withheld, and the next presentation asks again', async () => {
		const fixture = await open();
		request(fixture, 'close');
		await settle(fixture, () => bar(fixture) === null);
		expect(bar(fixture)).toBeNull();
		expect(drawn(fixture)).toBeFalsy();
		fixture.componentRef.setInput('content', await linkedContentDeck(PptxHandler, JSZip));
		await settle(fixture, () => bar(fixture) !== null);
		expect(bar(fixture)).not.toBeNull();
		expect(drawn(fixture)).toBeFalsy();
	});

	it('follows the host input: false withholds without a bar, true draws without asking', async () => {
		const fixture = await open({ allowExternalContent: false });
		expect(bar(fixture)).toBeNull();
		expect(drawn(fixture)).toBeFalsy();
		fixture.componentRef.setInput('allowExternalContent', true);
		await settle(fixture, () => drawn(fixture));
		expect(drawn(fixture)).toBeTruthy();
		expect(bar(fixture)).toBeNull();
	});

	it('shows the bar without the toolbar, and enables from a view-only viewer', async () => {
		const fixture = await open({ showToolbar: false });
		expect(bar(fixture)).not.toBeNull();
		request(fixture, 'action');
		await settle(fixture, () => drawn(fixture));
		expect(drawn(fixture)).toBeTruthy();
	});
});
