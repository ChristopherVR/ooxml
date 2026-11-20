// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { moveRegionFocus, regionEntryStop, regionKeyDirection } from './region-focus';

afterEach(() => document.body.replaceChildren());

const window_ = (): HTMLElement => {
	document.body.innerHTML = `
		<div class="window">
			<div class="ribbon"><button>Home</button><button aria-selected="true">Insert</button></div>
			<aside class="pane" hidden><button>Closed</button></aside>
			<div class="document" tabindex="0"><button>Shape</button></div>
			<div class="empty"></div>
			<footer class="status"><button>Zoom</button></footer>
		</div>`;
	return document.body.firstElementChild as HTMLElement;
};
const REGIONS = [
	{ selector: '.ribbon' },
	{ selector: '.pane' },
	{ selector: '.document', focus: (document_: HTMLElement) => (document_.focus(), true) },
	{ selector: '.empty' },
	{ selector: '.status' },
];
/** What has the keyboard: a region by its class, a stop by its text. */
const label = () =>
	document.activeElement?.className || document.activeElement?.textContent?.trim();

describe('the shared F6 region service', () => {
	it('moves to the next region and comes back round to the first', () => {
		const scope = window_();
		scope.querySelector<HTMLElement>('.document')!.focus();
		expect(moveRegionFocus(scope, REGIONS, 1)).toBe(true);
		expect(label()).toBe('empty');
		expect(moveRegionFocus(scope, REGIONS, 1)).toBe(true);
		expect(label()).toBe('Zoom');
		expect(moveRegionFocus(scope, REGIONS, 1)).toBe(true);
		// The current item of a region takes the keyboard, not its first stop.
		expect(label()).toBe('Insert');
		expect(moveRegionFocus(scope, REGIONS, 1)).toBe(true);
		expect(label()).toBe('document');
	});

	it('moves back the other way with Shift+F6', () => {
		const scope = window_();
		scope.querySelector<HTMLElement>('.document')!.focus();
		expect(moveRegionFocus(scope, REGIONS, -1)).toBe(true);
		expect(label()).toBe('Insert');
		expect(moveRegionFocus(scope, REGIONS, -1)).toBe(true);
		expect(label()).toBe('Zoom');
	});

	it('passes over a closed region and starts at the first with the keyboard outside', () => {
		const scope = window_();
		expect(moveRegionFocus(scope, REGIONS, 1)).toBe(true);
		expect(label()).toBe('Insert');
		expect(document.querySelector('.pane')!.contains(document.activeElement)).toBe(false);
	});

	it('takes the stop a region names over its first one', () => {
		document.body.innerHTML = `<form class="composer"><button>Attach</button><textarea></textarea></form>`;
		const composer = document.body.firstElementChild as HTMLElement;
		expect(regionEntryStop(composer).localName).toBe('button');
		expect(regionEntryStop(composer, 'textarea').localName).toBe('textarea');
	});

	it('gives a region with nothing to focus the keyboard itself', () => {
		const scope = window_();
		const empty = scope.querySelector<HTMLElement>('.empty')!;
		expect(regionEntryStop(empty)).toBe(empty);
		expect(empty.tabIndex).toBe(-1);
	});

	it('says there is nowhere to go when the window shows no region', () => {
		expect(moveRegionFocus(window_(), [{ selector: '.missing' }], 1)).toBe(false);
	});

	it('reads F6 and Shift+F6 and nothing else as the region key', () => {
		const key = (init: KeyboardEventInit) => regionKeyDirection(new KeyboardEvent('keydown', init));
		expect(key({ key: 'F6' })).toBe(1);
		expect(key({ key: 'F6', shiftKey: true })).toBe(-1);
		expect(key({ key: 'F6', ctrlKey: true })).toBeUndefined();
		expect(key({ key: 'F6', altKey: true })).toBeUndefined();
		expect(key({ key: 'F7' })).toBeUndefined();
	});
});
