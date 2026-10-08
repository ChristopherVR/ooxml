// @vitest-environment jsdom
import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import { registerOfficeUi } from '../index';
import type { OfficeUiRuler, RulerMarkerEventDetail } from './ruler';

beforeAll(() => registerOfficeUi());
afterEach(() => document.body.replaceChildren());

function ruler(setup: (el: OfficeUiRuler) => void = () => {}): OfficeUiRuler {
	const el = document.createElement('office-ui-ruler') as OfficeUiRuler;
	setup(el);
	document.body.append(el);
	return el;
}

describe('office-ui-ruler page features', () => {
	it('sizes itself, shades the margins and follows the zoom', () => {
		const el = ruler((r) => {
			r.extent = 816;
			r.marginStart = 96;
			r.marginEnd = 48;
			r.zoom = 1.5;
		});
		expect(el.style.width).toBe('816px');
		expect(el.style.getPropertyValue('zoom')).toBe('1.5');
		const bands = el.shadowRoot!.querySelectorAll<HTMLElement>('.margin');
		expect([...bands].map((b) => b.dataset.margin)).toEqual(['start', 'end']);
		expect(bands[0]!.style.width).toBe('96px');
		el.zoom = 1;
		expect(el.style.getPropertyValue('zoom')).toBe('');
	});

	it('is an image with a name when labelled and hidden otherwise', () => {
		const labelled = ruler((r) => (r.label = 'Ruler'));
		expect(labelled.getAttribute('role')).toBe('img');
		expect(labelled.getAttribute('aria-label')).toBe('Ruler');
		expect(labelled.hasAttribute('aria-hidden')).toBe(false);
		labelled.label = '';
		expect(labelled.getAttribute('aria-hidden')).toBe('true');
		expect(labelled.hasAttribute('role')).toBe(false);
	});

	it('positions markers and re-positions them when the property changes', () => {
		const el = ruler((r) => {
			r.markers = [
				{ name: 'left', position: 144 },
				{ name: 'first-line', position: 96, edge: 'top' },
			];
		});
		const left = () => el.shadowRoot!.querySelector<HTMLElement>('[data-marker="left"]')!;
		expect(left().style.left).toBe('144px');
		expect(left().classList.contains('bottom')).toBe(true);
		expect(
			el.shadowRoot!.querySelector('[data-marker="first-line"]')!.classList.contains('top'),
		).toBe(true);
		el.markers = [{ name: 'left', position: 192 }];
		expect(left().style.left).toBe('192px');
	});

	it('reports a dragged marker as move, then commit, then cancel', () => {
		const el = ruler((r) => (r.markers = [{ name: 'left', position: 96 }]));
		const seen: Array<[string, RulerMarkerEventDetail]> = [];
		for (const type of ['ruler-marker-move', 'ruler-marker-commit', 'ruler-marker-cancel'])
			el.addEventListener(type, (event) =>
				seen.push([type, (event as CustomEvent<RulerMarkerEventDetail>).detail]),
			);
		const marker = el.shadowRoot!.querySelector<HTMLElement>('[data-marker="left"]')!;
		const fire = (type: string, clientX: number) =>
			marker.dispatchEvent(new MouseEvent(type, { clientX, bubbles: true, button: 0 }));
		fire('pointerdown', 96);
		fire('pointermove', 144);
		fire('pointerup', 144);
		fire('pointerdown', 96);
		fire('pointercancel', 100);
		expect(seen).toEqual([
			['ruler-marker-move', { name: 'left', position: 144 }],
			['ruler-marker-commit', { name: 'left', position: 144 }],
			['ruler-marker-cancel', { name: 'left', position: 100 }],
		]);
	});

	it('removes the zoom from the reported position', () => {
		const el = ruler((r) => {
			r.zoom = 2;
			r.markers = [{ name: 'left', position: 96 }];
		});
		let position = 0;
		el.addEventListener('ruler-marker-commit', (event) => {
			position = (event as CustomEvent<RulerMarkerEventDetail>).detail.position;
		});
		const marker = el.shadowRoot!.querySelector<HTMLElement>('[data-marker="left"]')!;
		const fire = (type: string, clientX: number) =>
			marker.dispatchEvent(new MouseEvent(type, { clientX, bubbles: true, button: 0 }));
		fire('pointerdown', 0);
		fire('pointerup', 192);
		expect(position).toBe(96);
	});
});
