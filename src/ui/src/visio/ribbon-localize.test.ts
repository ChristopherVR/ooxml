import { afterEach, describe, expect, it } from 'vitest';
import { registerOfficeUi } from '../index';
import { localizeRibbon } from './ribbon-localize';

afterEach(() => document.body.replaceChildren());

type Gallery = HTMLElement & { state: unknown };

function ribbonWithGallery(): { ribbon: HTMLElement; gallery: Gallery } {
	registerOfficeUi();
	const ribbon = document.createElement('div');
	const gallery = document.createElement('office-ui-gallery') as Gallery;
	gallery.setAttribute('mode', 'panel');
	const state = (label: string) => ({
		id: 'styles',
		label: 'Styles',
		sections: [{ columns: 1, tileWidth: 40, tileHeight: 20, items: [{ id: 'one', label }] }],
	});
	gallery.state = state('Cut');
	ribbon.append(gallery);
	document.body.append(ribbon);
	return { ribbon, gallery };
}

describe('localizeRibbon', () => {
	it('changes language after a gallery has drawn without breaking its next update', () => {
		const { ribbon, gallery } = ribbonWithGallery();
		expect(() => {
			localizeRibbon(ribbon, 'de');
			localizeRibbon(ribbon, 'fr');
			// Lit draws into the gallery itself; a redraw after translation must still work.
			gallery.state = {
				id: 'styles',
				label: 'Styles',
				sections: [
					{ columns: 1, tileWidth: 40, tileHeight: 20, items: [{ id: 'one', label: 'Copy' }] },
				],
			};
		}).not.toThrow();
		expect(gallery.querySelector('[data-gallery-item="one"]')!.textContent).toContain('Copy');
	});

	it('still translates the labels and titles of ordinary ribbon elements', () => {
		registerOfficeUi();
		const ribbon = document.createElement('div');
		const button = document.createElement('office-ui-button');
		button.setAttribute('label', 'Insert');
		button.setAttribute('title', 'Insert');
		ribbon.append(button);
		document.body.append(ribbon);
		localizeRibbon(ribbon, 'de');
		expect(button.getAttribute('label')).toBe('Einfügen');
		localizeRibbon(ribbon, 'en');
		expect(button.getAttribute('label')).toBe('Insert');
	});
});
