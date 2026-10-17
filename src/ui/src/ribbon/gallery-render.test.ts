import { registerOfficeUi } from '../index';
import type { OfficeGalleryState } from './gallery';

beforeAll(() => registerOfficeUi());
afterEach(() => document.body.replaceChildren());

type Gallery = HTMLElement & { state: OfficeGalleryState | undefined };

const svg = (fill: string) =>
	`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10"><rect width="10" height="10" fill="${fill}"/></svg>`;
const state = (
	items: { id: string; fill: string; applied?: boolean }[],
	inline = false,
): OfficeGalleryState => ({
	id: 'styles',
	label: 'Styles',
	...(inline ? { inline: items.slice(0, 1).map(tileOf) } : {}),
	sections: [
		{ title: 'Theme', columns: 2, tileWidth: 20, tileHeight: 20, items: items.map(tileOf) },
	],
});
const tileOf = (item: { id: string; fill: string; applied?: boolean }) => ({
	id: item.id,
	label: item.id,
	applied: item.applied === true,
	preview: svg(item.fill),
});

function mount(initial: OfficeGalleryState): Gallery {
	const gallery = document.createElement('office-ui-gallery') as Gallery;
	gallery.setAttribute('mode', 'panel');
	gallery.state = initial;
	document.body.append(gallery);
	return gallery;
}
const graphic = (gallery: Gallery, id: string) =>
	gallery.querySelector(`[data-gallery-item="${id}"] svg`);

describe('office-ui-gallery redraws', () => {
	const items = [
		{ id: 'a', fill: '#111111' },
		{ id: 'b', fill: '#222222' },
	];

	it('leaves every tile alone when the host sets an equal state again', () => {
		const gallery = mount(state(items));
		const before = [graphic(gallery, 'a'), graphic(gallery, 'b')];
		expect(before.every(Boolean)).toBe(true);
		gallery.state = state(items);
		gallery.state = state(items);
		expect(graphic(gallery, 'a')).toBe(before[0]);
		expect(graphic(gallery, 'b')).toBe(before[1]);
	});

	it('keeps the graphics of tiles whose preview did not change when another tile does', () => {
		const gallery = mount(state(items));
		const a = graphic(gallery, 'a');
		const b = graphic(gallery, 'b');
		// The applied mark moves from nothing to a, and b's picture changes.
		gallery.state = state([
			{ id: 'a', fill: '#111111', applied: true },
			{ id: 'b', fill: '#333333' },
		]);
		expect(gallery.querySelector('[data-gallery-item="a"]')!.getAttribute('aria-pressed')).toBe(
			'true',
		);
		expect(graphic(gallery, 'a')).toBe(a);
		expect(graphic(gallery, 'b')).not.toBe(b);
		expect(graphic(gallery, 'b')!.innerHTML).toContain('#333333');
	});

	it('keeps the inline strip and the open popup each with their own graphic', () => {
		const gallery = mount(state(items, true));
		const tiles = gallery.querySelectorAll('[data-gallery-item="a"] svg');
		// One node cannot sit in two places, so the strip and the section never share one.
		expect(new Set(tiles).size).toBe(tiles.length);
		gallery.state = state(items, true);
		expect([...gallery.querySelectorAll('[data-gallery-item="a"] svg')]).toEqual([...tiles]);
	});

	it('draws a tile that is added later and forgets one that is removed', () => {
		const gallery = mount(state(items));
		gallery.state = state([...items, { id: 'c', fill: '#444444' }]);
		expect(graphic(gallery, 'c')).not.toBeNull();
		gallery.state = state([items[0]!]);
		expect(gallery.querySelector('[data-gallery-item="b"]')).toBeNull();
		gallery.state = state(items);
		// Back again: drawn afresh from its markup.
		expect(graphic(gallery, 'b')!.innerHTML).toContain('#222222');
	});

	it('keeps the trigger icon as one node however often the gallery draws', () => {
		const gallery = document.createElement('office-ui-gallery') as Gallery;
		gallery.setAttribute('icon', 'bold');
		gallery.setAttribute('label', 'Styles');
		gallery.state = state(items);
		document.body.append(gallery);
		const icon = gallery.querySelector('.trigger svg');
		expect(icon).not.toBeNull();
		// Hosts rewrite attributes and set state on every render of their own.
		for (let i = 0; i < 3; i++) {
			gallery.setAttribute('title', `Styles ${i}`);
			gallery.toggleAttribute('disabled', i % 2 === 0);
			gallery.state = state(items);
		}
		gallery.state = state([{ id: 'a', fill: '#555555', applied: true }, items[1]!]);
		expect(gallery.querySelector('.trigger svg')).toBe(icon);
	});

	it('still draws a first state', () => {
		const gallery = mount(state(items));
		expect(gallery.querySelectorAll('[data-gallery-item]')).toHaveLength(2);
		expect(graphic(gallery, 'a')).not.toBeNull();
	});
});
