// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { galleryLook, menuGallery } from './ribbon-gallery-menu';

describe('layout galleries', () => {
	it('describes margins and paper in inches', () => {
		expect(galleryLook('margins', 'narrow').detail).toContain('Left: 0.5"');
		expect(galleryLook('margins', 'moderate').detail).toContain('Left: 0.75"');
		expect(galleryLook('size', 'letter').detail).toBe('8.5" x 11"');
	});

	it('shows each position thumbnail at its named corner', () => {
		const mark = (value: string) =>
			galleryLook('drawingPosition', value).thumb.querySelector('rect.thumb-edge-on');
		expect(mark('top-left')?.getAttribute('x')).toBe('5');
		expect(mark('top-left')?.getAttribute('y')).toBe('5');
		expect(mark('bottom-right')?.getAttribute('x')).toBe('21');
		expect(mark('bottom-right')?.getAttribute('y')).toBe('28');
	});

	it('opens thumbnails, then drives the real select and its change event', () => {
		const host = document.createElement('div');
		document.body.append(host);
		const menu = menuGallery(
			'Orientation',
			'orientation',
			'orientation',
			[
				['portrait', 'Portrait'],
				['landscape', 'Landscape'],
			],
			(value) => ({ type: 'page', key: 'orientation', value }),
		);
		host.append(menu);
		const select = menu.querySelector('select')!;
		let changed = '';
		select.addEventListener('change', () => (changed = select.value));
		menu.click();
		const items = document.querySelectorAll<HTMLButtonElement>('.ribbon-gallery-menu button');
		expect(items).toHaveLength(2);
		expect(items[0]!.querySelector('svg.gallery-thumb')).not.toBeNull();
		expect(items[0]!.getAttribute('aria-checked')).toBe('true');
		items[1]!.click();
		expect(changed).toBe('landscape');
		expect(document.querySelector('.ribbon-gallery-menu')).toBeNull();
		host.remove();
	});

	it('focuses a momentary gallery and returns focus to its control on Escape', () => {
		const gallery = menuGallery(
			'Position',
			'pageNumber',
			'drawingPosition',
			[
				['top-left', 'Top Left'],
				['bottom-right', 'Bottom Right'],
			],
			(value) => ({ type: 'arrange', key: 'position', value }),
			{ momentary: true },
		);
		document.body.append(gallery);
		try {
			const control = gallery.querySelector('select')!;
			gallery.click();
			const [first, second] = document.querySelectorAll<HTMLButtonElement>(
				'.ribbon-gallery-menu button',
			);
			expect(document.activeElement).toBe(first);
			first!.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
			expect(document.activeElement).toBe(second);
			second!.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
			expect(document.activeElement).toBe(control);
		} finally {
			gallery.remove();
		}
	});

	it('draws border and line-spacing thumbnails and runs a trailing command', () => {
		const solid = (value: string) =>
			galleryLook('borders', value).thumb.querySelectorAll('.thumb-edge-on').length;
		expect(solid('none')).toBe(0);
		expect(solid('bottom')).toBe(1);
		expect(solid('outside')).toBe(4);
		expect(solid('all')).toBe(6);
		expect(
			galleryLook('lineSpacing', 'auto:480').thumb.querySelectorAll('.thumb-line'),
		).toHaveLength(3);

		const host = document.createElement('div');
		document.body.append(host);
		const menu = menuGallery(
			'Line spacing',
			'lineSpacing',
			'lineSpacing',
			[['auto:240', '1.0 lines']],
			(value) => ({ type: 'paragraph', key: 'lineSpacing', value }),
			{
				compact: true,
				commands: [
					{ label: 'Line Spacing Options…', action: { type: 'formatDialog', kind: 'paragraph' } },
				],
			},
		);
		host.append(menu);
		const actions: unknown[] = [];
		menu.addEventListener('ribbon-action', (event) => actions.push((event as CustomEvent).detail));
		menu.click();
		document.querySelector<HTMLButtonElement>('.gallery-command')!.click();
		expect(actions).toEqual([{ type: 'formatDialog', kind: 'paragraph' }]);
		host.remove();
	});
});
