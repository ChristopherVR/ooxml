import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import type { OfficeStatusBarState, OfficeStatusButton } from '../controls';
import { registerOfficeUi } from '../index';
import css from './status-bar.css?raw';

// Adapted from pptx-viewer's status-bar tests (`packages/shared/src/web-components/status-bar.test.ts`).
beforeAll(() => registerOfficeUi());
afterEach(() => document.body.replaceChildren());

type Bar = HTMLElement & { state: OfficeStatusBarState | undefined };

const STATE: OfficeStatusBarState = {
	items: [
		{ id: 'page', text: 'Page 3 of 10', live: true },
		{ id: 'language', text: 'English', narrowHide: true },
		{ id: 'save', text: 'Saved', tone: 'idle', narrowHide: true },
	],
	toggles: [{ id: 'notes', icon: 'message', label: 'Toggle notes', text: 'Notes', pressed: true }],
	views: [
		{ id: 'normal', icon: 'rectangle', label: 'Normal', pressed: true },
		{ id: 'read', icon: 'grid', label: 'Reading', pressed: false },
	],
	zoom: { percent: 99.6, outLabel: 'Zoom out', fitLabel: 'Fit', inLabel: 'Zoom in' },
};

function make(state: OfficeStatusBarState | null = STATE): Bar {
	const bar = document.createElement('office-ui-status-bar') as Bar;
	if (state) bar.state = state;
	document.body.append(bar);
	return bar;
}
const button = (bar: Bar, name: string) =>
	bar.shadowRoot!.querySelector<HTMLButtonElement>(`button[aria-label="${name}"]`)!;

describe('office-ui-status-bar (controlled)', () => {
	it('renders texts, named buttons, pressed state and the zoom percentage', () => {
		const bar = make();
		expect(bar.hasAttribute('data-controlled')).toBe(true);
		const page = bar.shadowRoot!.querySelector('[data-item="page"]')!;
		expect(page.textContent).toBe('Page 3 of 10');
		expect(page.getAttribute('aria-live')).toBe('polite');
		expect(bar.shadowRoot!.querySelector('[data-item="save"]')!.classList).toContain('narrow-hide');
		expect(button(bar, 'Toggle notes').getAttribute('aria-pressed')).toBe('true');
		expect(button(bar, 'Toggle notes').textContent).toBe('Notes');
		expect(button(bar, 'Normal').title).toBe('Normal');
		expect(button(bar, 'Reading').getAttribute('aria-pressed')).toBe('false');
		expect(button(bar, 'Fit').textContent).toBe('100%');
		expect(button(bar, 'Zoom in').querySelector('path')).not.toBeNull();
	});

	it('hides absent parts and patches buttons in place', () => {
		const bar = make();
		const normal = button(bar, 'Normal');
		bar.state = {
			...STATE,
			views: [{ ...(STATE.views as OfficeStatusButton[])[0]!, hidden: true }],
		};
		expect(button(bar, 'Normal')).toBe(normal);
		expect(normal.hidden).toBe(true);
		expect(bar.shadowRoot!.querySelector<HTMLElement>('.views')!.hidden).toBe(true);
		bar.state = { ...STATE, zoom: { ...STATE.zoom!, hidden: true } };
		expect(button(bar, 'Zoom in').closest<HTMLElement>('.group')!.hidden).toBe(true);
		bar.state = { items: STATE.items! };
		const groups = bar.shadowRoot!.querySelectorAll<HTMLElement>('.group');
		expect([...groups].map((g) => g.hidden)).toEqual([true, true]);
		expect(bar.shadowRoot!.querySelector<HTMLElement>('.toggles')!.hidden).toBe(true);
	});

	it('keeps glyph nodes across updates so a pressed pointer still clicks', () => {
		const bar = make();
		const path = button(bar, 'Zoom in').querySelector('path');
		bar.state = { ...STATE, zoom: { ...STATE.zoom!, percent: 120 } };
		expect(button(bar, 'Zoom in').querySelector('path')).toBe(path);
		expect(button(bar, 'Normal').querySelector('path')).not.toBeNull();
	});

	it('emits one activation per click and keeps Enter and Space inside', () => {
		const bar = make();
		const seen = vi.fn();
		bar.addEventListener('office-status-activate', (e) => seen((e as CustomEvent).detail));
		bar.state = { ...STATE, zoom: { ...STATE.zoom!, percent: 150 } };
		expect(seen).not.toHaveBeenCalled();
		for (const [name, id] of [
			['Toggle notes', 'notes'],
			['Reading', 'read'],
			['Zoom out', 'zoomOut'],
			['Fit', 'zoomFit'],
			['Zoom in', 'zoomIn'],
		] as const) {
			button(bar, name).click();
			expect(seen).toHaveBeenLastCalledWith({ id });
		}
		const outer = vi.fn();
		document.addEventListener('keydown', outer);
		button(bar, 'Zoom in').dispatchEvent(
			new KeyboardEvent('keydown', { key: ' ', bubbles: true, composed: true }),
		);
		document.removeEventListener('keydown', outer);
		expect(outer).not.toHaveBeenCalled();
	});

	it('puts the summary slot after the spacer and before the toggles and views', () => {
		const bar = make();
		const stats = document.createElement('span');
		stats.slot = 'summary';
		stats.textContent = 'Sum: 6';
		bar.append(stats);
		const parts = [...bar.shadowRoot!.querySelector('.bar')!.children];
		const at = (match: (node: Element) => boolean) => parts.findIndex(match);
		const summary = at((node) => node.getAttribute('name') === 'summary');
		expect((parts[summary] as HTMLSlotElement).assignedElements()).toEqual([stats]);
		expect(summary).toBeGreaterThan(at((node) => node.classList.contains('spacer')));
		expect(summary).toBeLessThan(at((node) => node.classList.contains('toggles')));
		expect(summary).toBeLessThan(at((node) => node.classList.contains('views')));
	});

	it('stays a plain slotted container without state', () => {
		const bar = make(null);
		expect(bar.hasAttribute('data-controlled')).toBe(false);
		for (const part of ['.items', '.toggles', '.views', '.zoom', '.spacer'])
			expect(bar.shadowRoot!.querySelector<HTMLElement>(part)!.hidden, part).toBe(true);
		expect(bar.shadowRoot!.querySelector('slot[name="end"]')).not.toBeNull();
	});

	it('draws an unavailable view disabled and does not activate it', () => {
		const bar = make({
			...STATE,
			views: [
				{ id: 'normal', icon: 'rectangle', label: 'Normal', pressed: true },
				{ id: 'layout', icon: 'grid', label: 'Page Layout', disabled: true },
			],
		});
		const seen = vi.fn();
		bar.addEventListener('office-status-activate', (e) => seen((e as CustomEvent).detail));
		expect(button(bar, 'Page Layout').disabled).toBe(true);
		button(bar, 'Page Layout').click();
		expect(seen).not.toHaveBeenCalled();
	});

	it('lets a product subclass rename the activation event', () => {
		const Base = customElements.get('office-ui-status-bar') as unknown as { new (): Bar };
		class Product extends Base {
			static activateEvent = 'status-request';
		}
		customElements.define('product-status-bar', Product as unknown as CustomElementConstructor);
		const bar = document.createElement('product-status-bar') as Bar;
		bar.state = STATE;
		document.body.append(bar);
		const seen = vi.fn();
		bar.addEventListener('status-request', (e) => seen((e as CustomEvent).detail));
		button(bar, 'Normal').click();
		expect(seen).toHaveBeenCalledWith({ id: 'normal' });
	});

	it('fades both separators through one token', () => {
		const rules = css.match(/opacity: var\(--office-status-bar-separator-opacity\);/g) ?? [];
		expect(rules).toHaveLength(2);
	});
});
