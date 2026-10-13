import { NgStyle } from '@angular/common';
import { NO_ERRORS_SCHEMA, signal } from '@angular/core';
import { TestBed } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import type { PptxElement } from 'pptx-viewer-core';
import { afterEach, beforeAll, expect, test } from 'vitest';

import {
	resolveViewerComponentResources,
	readViewerTestResource,
} from './component-resources.test-support';
import { ElementRendererShapeComponent } from './element-renderer-shape.component';

beforeAll(async () => {
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	await resolveViewerComponentResources();
});
afterEach(() => TestBed.resetTestingModule());

test('keeps its text component outside the hidden background layer and restores static styling', () => {
	TestBed.configureTestingModule({ imports: [ElementRendererShapeComponent] });
	TestBed.overrideComponent(ElementRendererShapeComponent, {
		set: {
			imports: [NgStyle],
			schemas: [NO_ERRORS_SCHEMA],
			template: readViewerTestResource('./element-renderer-shape.component.html'),
			templateUrl: undefined,
		},
	});
	const fixture = TestBed.createComponent(ElementRendererShapeComponent);
	const states = signal(
		new Map([['s::pptx-bg', { visible: false, cssAnimation: 'fade 1s both' }]]),
	);
	Object.assign(fixture.componentInstance, {
		element: signal({
			type: 'shape',
			id: 's',
			x: 10,
			y: 20,
			width: 200,
			height: 100,
			shapeType: 'rect',
			text: 'Persistent text',
			textSegments: [{ text: 'Persistent text', style: {} }],
		} as PptxElement),
		shapeContainerStyle: signal({ width: '200px', height: '100px', backgroundColor: 'red' }),
		subElementAnimStates: states,
	});
	fixture.detectChanges();
	const root = (fixture.nativeElement as HTMLElement).querySelector<HTMLElement>('.pptx-ng-shape')!;
	const paint = root.querySelector<HTMLElement>('[data-pptx-animation-layer="background"]')!;
	const text = root.querySelector('pptx-slide-text-block');
	expect(text).not.toBeNull();
	expect(paint.contains(text)).toBeFalsy();
	expect(paint.style.visibility).toBe('hidden');
	expect(paint.style.animation).toContain('fade');
	expect(root.style.backgroundColor).toBe('');
	states.set(new Map());
	fixture.detectChanges();
	expect(root.querySelector('[data-pptx-animation-layer="background"]')).toBeNull();
	expect(root.style.backgroundColor).toBe('red');
});
