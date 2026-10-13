/**
 * A picture's own frame (`a:ln`) and outer shadow, rendered through the real
 * `ImageRendererComponent` template. Angular painted neither: the picture box
 * clipped everything and no outline was mounted for a picture.
 */
import { NgStyle, NgTemplateOutlet } from '@angular/common';
import { NO_ERRORS_SCHEMA, signal } from '@angular/core';
import { TestBed } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import type { PptxElement } from 'pptx-viewer-core';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';

import { resolveViewerComponentResources } from './component-resources.test-support';
import { ImageRendererComponent } from './image-renderer.component';
import { StrokeOutlineComponent } from './stroke-outline.component';

beforeAll(async () => {
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	await resolveViewerComponentResources();
});
afterEach(() => TestBed.resetTestingModule());

function picture(withFrame: boolean): PptxElement {
	return {
		type: 'picture',
		id: 'pic-frame',
		x: 0,
		y: 0,
		width: 200,
		height: 150,
		imageData: 'data:image/png;base64,AA==',
		...(withFrame
			? {
					shapeStyle: {
						strokeWidth: 10.67,
						strokeFillMode: 'solid',
						strokeColor: '#FF0000',
						shadowColor: '#000000',
						shadowOpacity: 0.6,
						shadowBlur: 5.33,
						shadowAngle: 45,
						shadowDistance: 10.67,
					},
				}
			: {}),
	} as unknown as PptxElement;
}

function renderBox(element: PptxElement): HTMLElement {
	TestBed.configureTestingModule({ imports: [ImageRendererComponent] });
	TestBed.overrideComponent(ImageRendererComponent, {
		set: {
			imports: [NgStyle, NgTemplateOutlet, StrokeOutlineComponent],
			schemas: [NO_ERRORS_SCHEMA],
		},
	});
	const fixture = TestBed.createComponent(ImageRendererComponent);
	// Plain JIT does not discover signal-input metadata, so hand the input a signal.
	Object.assign(fixture.componentInstance, { element: signal(element) });
	fixture.detectChanges();
	return (fixture.nativeElement as HTMLElement).querySelector('.pptx-ng-image') as HTMLElement;
}

describe('imageRendererComponent picture frame (a:ln + a:outerShdw)', () => {
	it('paints the whole frame and the shadow from an unclipped box', () => {
		const box = renderBox(picture(true));
		expect(box.style.overflow).toBe('visible');
		expect(box.style.filter).toMatch(/drop-shadow\(/);
		const surface = box.querySelector<HTMLElement>(':scope > .pptx-ng-image-surface');
		expect(surface?.style.overflow).toBe('hidden');
		expect(surface?.querySelector('img')).not.toBeNull();
		const path = box.querySelector(':scope > svg path');
		expect(path?.getAttribute('stroke')).toBe('#FF0000');
		expect(
			box!.querySelector('img')!.compareDocumentPosition(path!) & Node.DOCUMENT_POSITION_FOLLOWING,
		).toBeTruthy();
		expect(Number(path?.getAttribute('stroke-width'))).toBeCloseTo(10.67);
	});

	it('keeps the single clipped box for a picture without a frame', () => {
		const box = renderBox(picture(false));
		expect(box.style.overflow).toBe('hidden');
		expect(box.querySelector('.pptx-ng-image-surface')).toBeNull();
		expect(box.querySelector(':scope > img')).not.toBeNull();
	});
});
