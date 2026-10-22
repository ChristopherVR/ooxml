import { signal } from '@angular/core';
import { TestBed } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { provideTranslateService } from '@ngx-translate/core';
import type { PptxElement } from 'pptx-viewer-core';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';

import { registerPptxWebControls } from '../../../../../../src/ui/src/pptx/web-components';
import { TextWarpGalleryComponent } from './text-warp-gallery.component';

beforeAll(() => {
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	registerPptxWebControls();
});
afterEach(() => TestBed.resetTestingModule());

const text = (textWarpPreset?: string) =>
	({
		id: 'title',
		type: 'text',
		x: 0,
		y: 0,
		width: 200,
		height: 60,
		text: 'Hello',
		textStyle: { fontSize: 24, ...(textWarpPreset ? { textWarpPreset } : {}) },
	}) as PptxElement;

function mount(element: PptxElement) {
	TestBed.configureTestingModule({
		imports: [TextWarpGalleryComponent],
		providers: [provideTranslateService({ fallbackLang: 'en' })],
	});
	// The Vitest JIT build does not wire signal inputs, so declare them and hand each
	// field a plain signal.
	TestBed.overrideComponent(TextWarpGalleryComponent, { add: { inputs: ['element'] } });
	const fixture = TestBed.createComponent(TextWarpGalleryComponent);
	fixture.componentRef.setInput('element', signal(element));
	fixture.detectChanges();
	const host = fixture.nativeElement as HTMLElement;
	return { fixture, host };
}

const pressed = (host: HTMLElement) =>
	[...host.querySelectorAll('[data-gallery-item][aria-pressed="true"]')].map((el) =>
		el.getAttribute('data-gallery-item'),
	);

describe('textWarpGalleryComponent', () => {
	it('draws every preset on the shared panel and marks the current one', () => {
		const { host } = mount(text('textArchUp'));
		const gallery = host.querySelector('pptx-ui-ribbon-gallery')!;
		expect(gallery.getAttribute('mode')).toBe('panel');
		expect(gallery.hasAttribute('compact')).toBeTruthy();
		expect(host.querySelectorAll('[data-gallery-item]')).toHaveLength(39);
		expect(pressed(host)).toStrictEqual(['textArchUp']);
	});

	it('emits the text style patch for a pick, and No Transform clears the warp', () => {
		const { fixture, host } = mount(text());
		expect(pressed(host)).toStrictEqual(['textNoShape']);
		const patches: Partial<PptxElement>[] = [];
		fixture.componentInstance.patch.subscribe((patch) => patches.push(patch));
		host.querySelector<HTMLButtonElement>('[data-gallery-item="textWave1"]')!.click();
		host.querySelector<HTMLButtonElement>('[data-gallery-item="textNoShape"]')!.click();
		expect(patches).toStrictEqual([
			{ textStyle: { fontSize: 24, textWarpPreset: 'textWave1' } },
			{ textStyle: { fontSize: 24, textWarpPreset: undefined } },
		]);
	});
});
