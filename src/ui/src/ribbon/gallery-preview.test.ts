import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest';
import type { OfficeGalleryState } from '../controls';
import { registerOfficeUi } from '../index';

beforeAll(() => registerOfficeUi());
beforeEach(() => vi.useFakeTimers());
afterEach(() => {
	vi.useRealTimers();
	document.body.replaceChildren();
});

type Gallery = HTMLElement & {
	state: OfficeGalleryState | undefined;
	readonly trigger: HTMLButtonElement;
	open: boolean;
	close(): void;
};

const preview =
	'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 40 30"><rect width="40" height="30"/></svg>';
const actions = [
	{ id: 'all', label: 'Apply to All Slides' },
	{ id: 'selected', label: 'Apply to Selected Slides' },
];
const STATE: OfficeGalleryState = {
	id: 'themeBackgrounds',
	label: 'Background Styles',
	preview: true,
	footer: [{ id: 'formatBackground', label: 'Format Background...' }],
	sections: [
		{
			columns: 2,
			tileWidth: 40,
			tileHeight: 30,
			items: [
				{ id: 'one', label: 'Style 1', preview, actions },
				{ id: 'two', label: 'Style 2', preview, actions },
			],
		},
	],
};

function mount(state: OfficeGalleryState = STATE, inline = false): Gallery {
	const gallery = document.createElement('office-ui-gallery') as Gallery;
	if (inline) gallery.setAttribute('mode', 'inline');
	gallery.state = state;
	document.body.append(gallery);
	return gallery;
}

function tile(gallery: Gallery, id: string): HTMLButtonElement {
	return gallery.querySelector<HTMLButtonElement>(`[data-gallery-item="${id}"]`)!;
}

function hover(element: Element): void {
	element.dispatchEvent(new PointerEvent('pointerenter', { pointerType: 'mouse' }));
}

function leave(element: Element): void {
	element.dispatchEvent(new PointerEvent('pointerleave', { pointerType: 'mouse' }));
}

describe('office-ui-gallery preview', () => {
	it('previews a hovered tile after a pause and ends it when the pointer leaves', () => {
		const gallery = mount();
		gallery.open = true;
		const started = vi.fn();
		const ended = vi.fn();
		gallery.addEventListener('office-gallery-preview', (e) => started((e as CustomEvent).detail));
		gallery.addEventListener('office-gallery-preview-end', (e) => ended((e as CustomEvent).detail));
		hover(tile(gallery, 'one'));
		expect(started).not.toHaveBeenCalled();
		vi.advanceTimersByTime(130);
		expect(started).toHaveBeenCalledWith({ gallery: 'themeBackgrounds', itemId: 'one' });
		leave(tile(gallery, 'one'));
		expect(ended).not.toHaveBeenCalled();
		vi.advanceTimersByTime(70);
		expect(ended).toHaveBeenCalledWith({ gallery: 'themeBackgrounds' });
	});

	it('moves between tiles without ending the preview', () => {
		const gallery = mount();
		gallery.open = true;
		const events: string[] = [];
		gallery.addEventListener('office-gallery-preview', (e) =>
			events.push(`start ${(e as CustomEvent).detail.itemId}`),
		);
		gallery.addEventListener('office-gallery-preview-end', () => events.push('end'));
		hover(tile(gallery, 'one'));
		vi.advanceTimersByTime(130);
		leave(tile(gallery, 'one'));
		hover(tile(gallery, 'two'));
		vi.advanceTimersByTime(200);
		expect(events).toEqual(['start one', 'start two']);
	});

	it('ends the preview after a pick and when the popup closes', () => {
		const gallery = mount();
		gallery.open = true;
		const order: string[] = [];
		gallery.addEventListener('office-gallery-preview-end', () => order.push('end'));
		gallery.addEventListener('office-gallery-pick', () => order.push('pick'));
		hover(tile(gallery, 'one'));
		vi.advanceTimersByTime(130);
		tile(gallery, 'one').click();
		expect(order).toEqual(['pick', 'end']);
		gallery.open = true;
		hover(tile(gallery, 'two'));
		vi.advanceTimersByTime(130);
		gallery.close();
		expect(order).toEqual(['pick', 'end', 'end']);
	});

	it('previews inline tiles and ignores touch pointers', () => {
		const gallery = mount(STATE, true);
		const started = vi.fn();
		gallery.addEventListener('office-gallery-preview', started);
		tile(gallery, 'one').dispatchEvent(new PointerEvent('pointerenter', { pointerType: 'touch' }));
		vi.advanceTimersByTime(200);
		expect(started).not.toHaveBeenCalled();
		hover(tile(gallery, 'one'));
		vi.advanceTimersByTime(200);
		expect(started).toHaveBeenCalledOnce();
	});

	it('does not preview unless the state asks for it', () => {
		const gallery = mount({ ...STATE, preview: false });
		gallery.open = true;
		const started = vi.fn();
		gallery.addEventListener('office-gallery-preview', started);
		hover(tile(gallery, 'one'));
		vi.advanceTimersByTime(200);
		expect(started).not.toHaveBeenCalled();
	});
});

describe('office-ui-gallery tile menu and footer', () => {
	it('opens a context menu on a tile and picks its action', () => {
		const gallery = mount();
		gallery.open = true;
		const picks: unknown[] = [];
		gallery.addEventListener('office-gallery-pick', (e) => picks.push((e as CustomEvent).detail));
		tile(gallery, 'two').dispatchEvent(
			new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 30, clientY: 40 }),
		);
		const items = gallery.querySelectorAll<HTMLButtonElement>('.tile-menu [role="menuitem"]');
		expect([...items].map((item) => item.textContent)).toEqual([
			'Apply to All Slides',
			'Apply to Selected Slides',
		]);
		items[1]!.click();
		expect(picks).toEqual([{ gallery: 'themeBackgrounds', itemId: 'two', action: 'selected' }]);
		expect(gallery.querySelector('.tile-menu')).toBeNull();
	});

	it('opens the menu from the keyboard and closes it on Escape', () => {
		const gallery = mount();
		gallery.open = true;
		tile(gallery, 'one').dispatchEvent(
			new KeyboardEvent('keydown', { key: 'F10', shiftKey: true, bubbles: true, cancelable: true }),
		);
		const menu = gallery.querySelector<HTMLElement>('.tile-menu')!;
		expect(menu).not.toBeNull();
		menu.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
		expect(gallery.querySelector('.tile-menu')).toBeNull();
		expect(gallery.open).toBe(true);
	});

	it('a plain click picks the tile without an action, and the footer picks its command', () => {
		const gallery = mount();
		gallery.open = true;
		const picks: unknown[] = [];
		gallery.addEventListener('office-gallery-pick', (e) => picks.push((e as CustomEvent).detail));
		tile(gallery, 'one').click();
		gallery.open = true;
		gallery.querySelector<HTMLButtonElement>('[data-gallery-footer="formatBackground"]')!.click();
		expect(picks).toEqual([
			{ gallery: 'themeBackgrounds', itemId: 'one' },
			{ gallery: 'themeBackgrounds', itemId: 'formatBackground' },
		]);
	});
});
