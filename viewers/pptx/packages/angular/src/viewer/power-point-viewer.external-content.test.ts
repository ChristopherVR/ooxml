/**
 * PowerPoint's security bar in the Angular viewer: Enable Content reloads the presentation as it
 * stands, so an edit saved before it stays (and the next save writes it), an unsaved edit stays
 * unsaved, and the bar shows without the toolbar too.
 */
import { TestBed } from '@angular/core/testing';
import type { ComponentFixture } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { provideTranslateService } from '@ngx-translate/core';
import JSZip from 'jszip';
import { PptxHandler } from 'pptx-viewer-core';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';

import { linkedContentDeck } from '../../../shared/test-utils/external-content-contract.mjs';
import { resolveViewerComponentResources } from './component-resources.test-support';
import { PowerPointViewerComponent } from './power-point-viewer.component';
import { declareSignalFields } from './signal-inputs.test-support';

const ORIGIN = 'https://linked-content.test';

beforeAll(async () => {
	declareSignalFields(PowerPointViewerComponent, import.meta.dirname);
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	await resolveViewerComponentResources();
});

afterEach(() => {
	TestBed.resetTestingModule();
	document.body.replaceChildren();
	localStorage.clear();
});

type Viewer = ComponentFixture<PowerPointViewerComponent>;

async function settle(fixture: Viewer, done: () => boolean): Promise<void> {
	for (let attempt = 0; attempt < 200 && !done(); attempt += 1) {
		fixture.detectChanges();
		await fixture.whenStable();
		await new Promise((resolve) => {
			setTimeout(resolve, 20);
		});
	}
	fixture.detectChanges();
}

const bar = (fixture: Viewer) =>
	(fixture.nativeElement as HTMLElement).querySelector('pptx-ui-external-content-bar');
const referencesOrigin = (fixture: Viewer) =>
	JSON.stringify(fixture.componentInstance.getSlides()).includes(ORIGIN);

async function open(inputs: Record<string, unknown>): Promise<Viewer> {
	TestBed.configureTestingModule({ providers: [provideTranslateService()] });
	const fixture = TestBed.createComponent(PowerPointViewerComponent);
	document.body.append(fixture.nativeElement as HTMLElement);
	fixture.componentRef.setInput('content', await linkedContentDeck(PptxHandler, JSZip));
	fixture.componentRef.setInput('googleFonts', false);
	for (const [name, value] of Object.entries(inputs)) {
		fixture.componentRef.setInput(name, value);
	}
	await settle(fixture, () => bar(fixture) !== null);
	return fixture;
}

function enable(fixture: Viewer): void {
	bar(fixture)!.dispatchEvent(
		new CustomEvent('external-content-request', { detail: { id: 'action' } }),
	);
}

const slideXml = async (bytes: Uint8Array) =>
	(await JSZip.loadAsync(bytes)).file('ppt/slides/slide1.xml')!.async('string');
const offset = (xml: string) => xml.match(/<a:off [^>]*>/u)?.[0];

describe('angular viewer Enable Content', { timeout: 60_000 }, () => {
	it('keeps edits saved before Enable Content, and the next save writes them', async () => {
		const fixture = await open({ canEditInput: true });
		const viewer = fixture.componentInstance;
		const [first] = viewer.getElements(0);
		viewer.updateElement(first!.id, { x: first!.x + 40 });
		fixture.detectChanges();
		expect(viewer.getElements(0)[0]!.x, 'edited').toBe(first!.x + 40);
		const saved = await viewer.getSaveSnapshot();
		expect(viewer.acknowledgeSave(saved.revision).status).toBe('clean');
		enable(fixture);
		await settle(fixture, () => referencesOrigin(fixture));
		expect(referencesOrigin(fixture)).toBe(true);
		expect(bar(fixture)).toBeNull();
		expect(viewer.getElements(0)[0]!.x).toBe(first!.x + 40);
		expect(viewer.isDirty()).toBe(false);
		const next = await viewer.getSaveSnapshot();
		expect(offset(await slideXml(next.content))).toBe(offset(await slideXml(saved.content)));
		expect(await slideXml(next.content)).toContain('r:link="rIdLinkedFill"');
	});

	it('keeps unsaved edits unsaved through Enable Content', async () => {
		const fixture = await open({ canEditInput: true });
		const viewer = fixture.componentInstance;
		const [first] = viewer.getElements(0);
		viewer.updateElement(first!.id, { x: first!.x + 40 });
		enable(fixture);
		await settle(fixture, () => referencesOrigin(fixture));
		expect(viewer.getElements(0)[0]!.x).toBe(first!.x + 40);
		expect(viewer.isDirty()).toBe(true);
	});

	it('shows the bar without the toolbar, and enables from a view-only viewer', async () => {
		const fixture = await open({ showToolbar: false });
		expect(bar(fixture)).not.toBeNull();
		enable(fixture);
		await settle(fixture, () => referencesOrigin(fixture));
		expect(referencesOrigin(fixture)).toBe(true);
	});
});
