/**
 * A dialog flag raised as a host request through a real session request store, read through the
 * binding's store adapter: the flag is handled once, even when an effect reading the store reacts
 * to the request (the loop React 18 hit with flag effects keyed on more than the flag).
 */
import { effect, signal } from '@angular/core';
import { TestBed } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { createBuildingBlockRequests } from 'ooxml-ui/pptx';
import { afterEach, beforeAll, expect, test, vi } from 'vitest';

import { controllerSignal } from '../viewer/controller-signal';
import { routeDialogFlag } from './route-dialog-flag';

beforeAll(() => {
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
});

afterEach(() => {
	TestBed.resetTestingModule();
});

const flush = () => TestBed.tick();

test('a dialog flag raises one request through the session store', () => {
	const onRequest = vi.fn();
	const requests = createBuildingBlockRequests({
		state: () => ({ available: true, needsHostUi: true }),
		onRequest,
	});
	const open = signal(false);
	let seen = 0;
	TestBed.runInInjectionContext(() => {
		const pending = controllerSignal(requests);
		// Something that reacts to the store, as a template or effect would.
		effect(() => {
			void pending();
			seen += 1;
		});
		routeDialogFlag(
			() => open(),
			() => {
				open.set(false);
				requests.request({ kind: 'smartArt' });
			},
		);
	});
	flush();
	open.set(true);
	flush();
	expect(onRequest).toHaveBeenCalledOnce();
	expect(open()).toBe(false);
	expect(requests.getSnapshot()).toEqual({ kind: 'smartArt' });
	expect(seen).toBeGreaterThan(1);
	open.set(true);
	flush();
	expect(onRequest).toHaveBeenCalledTimes(2);
});
