/**
 * The Angular wiring of the shared viewport navigation: the service follows
 * the `<main>` query (which renders only once a deck has loaded), reads its
 * accessors live and detaches with its injector.
 */
import { signal } from '@angular/core';
import { TestBed } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';

import { ViewerViewportNavigationService } from './viewer-viewport-navigation.service';
import { ViewerZoomService } from './viewer-zoom.service';

beforeAll(() => {
	try {
		TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	} catch {
		// Already initialised by another spec in this worker.
	}
});
afterEach(() => TestBed.resetTestingModule());

function ctrlWheel(node: HTMLElement): boolean {
	const event = new WheelEvent('wheel', {
		deltaY: -100,
		cancelable: true,
		clientX: 10,
		clientY: 10,
	});
	// Happy DOM's WheelEvent omits the inherited mouse modifier properties.
	Object.defineProperty(event, 'ctrlKey', { value: true });
	node.dispatchEvent(event);
	return event.defaultPrevented;
}

function scrollingMain(): { node: HTMLElement; drag: () => number } {
	const node = document.createElement('main');
	node.innerHTML = '<div aria-roledescription="slide"></div>';
	let scroll = 0;
	Object.defineProperties(node, {
		clientWidth: { configurable: true, get: () => 100 },
		clientHeight: { configurable: true, get: () => 100 },
		scrollWidth: { configurable: true, get: () => 300 },
		scrollHeight: { configurable: true, get: () => 300 },
		scrollLeft: {
			configurable: true,
			get: () => scroll,
			set: (value: number) => {
				scroll = Math.max(0, value);
			},
		},
		scrollTop: { configurable: true, get: () => 0, set: () => {} },
	});
	vi.spyOn(node, 'setPointerCapture').mockImplementation(() => {});
	vi.spyOn(node, 'releasePointerCapture').mockImplementation(() => {});
	const drag = (): number => {
		scroll = 100;
		const pointer = (type: string, clientX: number) => {
			const event = new MouseEvent(type, {
				bubbles: true,
				cancelable: true,
				button: 0,
				clientX,
				clientY: 10,
			});
			Object.defineProperties(event, { pointerId: { value: 1 }, pointerType: { value: 'mouse' } });
			node.dispatchEvent(event);
		};
		pointer('pointerdown', 200);
		pointer('pointermove', 150);
		pointer('pointerup', 150);
		return scroll;
	};
	return { node, drag };
}

function setup() {
	TestBed.configureTestingModule({
		providers: [ViewerZoomService, ViewerViewportNavigationService],
	});
	const zoom = TestBed.inject(ViewerZoomService);
	const main = signal<HTMLElement | undefined>(undefined);
	const dragPans = signal(true);
	const presenting = signal(false);
	TestBed.inject(ViewerViewportNavigationService).setup(() => main(), {
		dragPans: () => dragPans(),
		presenting: () => presenting(),
	});
	TestBed.tick();
	return { zoom, main, dragPans, presenting };
}

describe('viewer viewport navigation service', () => {
	it('attaches once <main> renders, and moves to a replacement', () => {
		const env = setup();
		const first = scrollingMain();
		env.main.set(first.node);
		TestBed.tick();
		expect(ctrlWheel(first.node)).toBeTruthy();
		expect(env.zoom.zoom()).toBeGreaterThan(1);

		const second = scrollingMain();
		env.main.set(second.node);
		TestBed.tick();
		expect(ctrlWheel(first.node)).toBeFalsy();
		expect(ctrlWheel(second.node)).toBeTruthy();
	});

	it('reads the drag and presenting accessors live', () => {
		const env = setup();
		const { node, drag } = scrollingMain();
		env.main.set(node);
		TestBed.tick();
		expect(drag()).toBe(150);
		env.dragPans.set(false);
		expect(drag()).toBe(100);
		env.presenting.set(true);
		expect(ctrlWheel(node)).toBeFalsy();
	});

	it('detaches when its injector is destroyed', () => {
		const env = setup();
		const { node } = scrollingMain();
		env.main.set(node);
		TestBed.tick();
		TestBed.resetTestingModule();
		expect(ctrlWheel(node)).toBeFalsy();
	});
});
