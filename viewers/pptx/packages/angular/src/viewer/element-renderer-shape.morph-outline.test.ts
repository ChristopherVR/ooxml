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
import { StrokeOutlineComponent } from './stroke-outline.component';

/**
 * A resizing shape's Morph re-paths its outline per frame (`morph-shape-carry`
 * in ooxml-ui), through a descendant rule on `[data-pptx-outline] path`. The
 * outline overlay has to carry that marker in every binding or the outline
 * scales with the box instead.
 */
beforeAll(async () => {
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	await resolveViewerComponentResources();
});
afterEach(() => TestBed.resetTestingModule());

function render(style: Record<string, unknown>): HTMLElement {
	TestBed.configureTestingModule({ imports: [ElementRendererShapeComponent] });
	TestBed.overrideComponent(ElementRendererShapeComponent, {
		set: {
			imports: [NgStyle, StrokeOutlineComponent],
			schemas: [NO_ERRORS_SCHEMA],
			template: readViewerTestResource('./element-renderer-shape.component.html'),
			templateUrl: undefined,
		},
	});
	const fixture = TestBed.createComponent(ElementRendererShapeComponent);
	Object.assign(fixture.componentInstance, {
		element: signal({
			type: 'shape',
			id: 's',
			x: 0,
			y: 0,
			width: 400,
			height: 160,
			shapeType: 'roundRect',
			shapeStyle: style,
		} as unknown as PptxElement),
		shapeContainerStyle: signal({ width: '400px', height: '160px' }),
		subElementAnimStates: signal(new Map()),
	});
	fixture.detectChanges();
	return fixture.nativeElement as HTMLElement;
}

test('paints each strand of a compound outline as a sibling path, in order', () => {
	// The compound carry addresses strands as `[data-pptx-outline] path:nth-of-type(n)`.
	const host = render({
		fillColor: '#C07020',
		strokeColor: '#FFFFFF',
		strokeWidth: 12,
		compoundLine: 'tri',
	});
	expect(host.querySelectorAll('svg[data-pptx-outline] > path')).toHaveLength(3);
});

test('stamps data-pptx-outline on the outline overlay svg', () => {
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
	Object.assign(fixture.componentInstance, {
		element: signal({
			type: 'shape',
			id: 's',
			x: 0,
			y: 0,
			width: 400,
			height: 160,
			shapeType: 'roundRect',
			shapeStyle: { fillColor: '#C07020', strokeColor: '#FFFFFF', strokeWidth: 8 },
		} as unknown as PptxElement),
		shapeContainerStyle: signal({ width: '400px', height: '160px' }),
		subElementAnimStates: signal(new Map()),
	});
	fixture.detectChanges();
	const outline = (fixture.nativeElement as HTMLElement).querySelector('svg[data-pptx-outline]');
	expect(outline).not.toBeNull();
	expect(outline?.classList.contains('pptx-ng-gradient-outline')).toBe(true);
});
