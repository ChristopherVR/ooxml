import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import type { OfficeGalleryState } from '../controls.js';
import { registerOfficeUi } from '../index.js';
import { parseSvgPreview } from './safe-svg.js';

// Adapted from pptx-viewer's ribbon-gallery tests (`packages/shared/src/web-components/ribbon-gallery.test.ts`).
beforeAll(() => registerOfficeUi());
afterEach(() => document.body.replaceChildren());

type Gallery = HTMLElement & {
	state: OfficeGalleryState | undefined;
	readonly trigger: HTMLButtonElement;
	readonly popup: HTMLElement;
	disabled: boolean;
	open: boolean;
};

const preview =
	'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 40 30"><rect width="40" height="30"/></svg>';
const STYLES: OfficeGalleryState = {
	id: 'shapeStyles',
	label: 'Shape Styles',
	moreLabel: 'More Shape Styles',
	sections: [
		{
			title: 'Theme Styles',
			columns: 3,
			tileWidth: 40,
			tileHeight: 30,
			items: Array.from({ length: 6 }, (_, i) => ({
				id: `style${i}`,
				label: `Style ${i}`,
				applied: i === 2,
				preview,
			})),
		},
	],
};

function mount(state: OfficeGalleryState = STYLES, inline = true): Gallery {
	const gallery = document.createElement('office-ui-gallery') as Gallery;
	if (inline) gallery.setAttribute('mode', 'inline');
	gallery.state = state;
	document.body.append(gallery);
	return gallery;
}

describe('office-ui-gallery', () => {
	it('navigates dropdown sections without applying a choice until activation', () => {
		const gallery = mount(
			{
				...STYLES,
				sections: [
					...STYLES.sections,
					{ ...STYLES.sections[0]!, items: [{ id: 'next', label: 'Next' }] },
				],
			},
			false,
		);
		const pick = vi.fn();
		gallery.addEventListener('office-gallery-pick', pick);
		gallery.trigger.dispatchEvent(
			new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }),
		);
		for (const [key, expected] of [
			['End', 'next'],
			['Home', 'style0'],
			['ArrowUp', 'style4'],
			['ArrowRight', 'style5'],
			['ArrowRight', 'next'],
		] as const) {
			document.activeElement!.dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true }));
			expect(document.activeElement?.getAttribute('data-gallery-item')).toBe(expected);
		}
		expect(pick).not.toHaveBeenCalled();
		(document.activeElement as HTMLButtonElement).click();
		expect(pick).toHaveBeenCalledOnce();
		expect(gallery.open).toBe(false);
		expect(document.activeElement).toBe(gallery.trigger);
	});
	it('reuses labelled tiles in a panel and keeps focus while the selection refreshes', () => {
		const gallery = mount();
		gallery.setAttribute('mode', 'panel');
		gallery.open = true;
		expect(gallery.open).toBe(false);
		expect(gallery.querySelector('.trigger')).toBeNull();
		expect(gallery.querySelector('.tile-label')?.textContent).toBe('Style 0');
		const picked = vi.fn();
		gallery.addEventListener('office-gallery-pick', (event) => {
			const detail = (event as CustomEvent<{ itemId: string }>).detail;
			picked(detail.itemId);
			gallery.state = {
				...STYLES,
				sections: STYLES.sections.map((section) => ({
					...section,
					items: section.items.map((item) => ({ ...item, applied: item.id === detail.itemId })),
				})),
			};
		});
		gallery.querySelector<HTMLButtonElement>('[data-gallery-item="style0"]')!.focus();
		for (const [key, expected] of [
			['ArrowDown', 'style3'],
			['End', 'style5'],
			['Home', 'style0'],
			['ArrowLeft', 'style5'],
		] as const) {
			document.activeElement!.dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true }));
			expect(document.activeElement?.getAttribute('data-gallery-item')).toBe(expected);
			expect(document.activeElement?.getAttribute('aria-pressed')).toBe('true');
		}
		expect(picked.mock.calls.flat()).toEqual(['style3', 'style5', 'style0', 'style5']);
		gallery.disabled = true;
		gallery.querySelector<HTMLButtonElement>('[data-gallery-item="style2"]')!.click();
		expect(picked).toHaveBeenCalledTimes(4);
	});
	it('stamps the popup hook on the closed placeholder', () => {
		const gallery = mount();
		expect(gallery.contains(gallery.popup)).toBe(false);
		expect(gallery.popup.getAttribute('data-gallery-popup')).toBe('shapeStyles');
	});

	it('renders inline previews and emits one pick from the popup', () => {
		const gallery = mount();
		const pick = vi.fn();
		gallery.addEventListener('office-gallery-pick', (e) => pick((e as CustomEvent).detail));
		expect(gallery.querySelectorAll('[data-gallery-item]')).toHaveLength(6);
		expect(
			gallery.querySelector('[data-gallery-item="style2"]')!.getAttribute('aria-pressed'),
		).toBe('true');
		expect(gallery.querySelector('.tile svg')).not.toBeNull();
		expect(gallery.trigger.title).toBe('More Shape Styles');
		gallery.trigger.click();
		expect(gallery.popup.querySelector('.heading')!.textContent).toBe('Theme Styles');
		gallery.popup.querySelector<HTMLButtonElement>('[data-gallery-item="style4"]')!.click();
		expect(pick).toHaveBeenCalledExactlyOnceWith({ gallery: 'shapeStyles', itemId: 'style4' });
		expect(gallery.open).toBe(false);
		expect(gallery.contains(gallery.popup)).toBe(false);
		expect(document.activeElement).toBe(gallery.trigger);
	});

	it('keeps a focused popup tile through a refresh and closes on Escape', () => {
		const gallery = mount();
		gallery.trigger.dispatchEvent(
			new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }),
		);
		const itemId = document.activeElement?.getAttribute('data-gallery-item');
		expect(itemId).toBe('style0');
		gallery.state = { ...STYLES };
		expect(gallery.popup.contains(document.activeElement)).toBe(true);
		expect(document.activeElement?.getAttribute('data-gallery-item')).toBe(itemId);
		document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
		expect(gallery.open).toBe(false);
		expect(document.activeElement).toBe(gallery.trigger);
	});

	it('closes when disabled, refuses picks and clears without state', () => {
		const gallery = mount();
		const pick = vi.fn();
		gallery.addEventListener('office-gallery-pick', pick);
		gallery.open = true;
		gallery.disabled = true;
		expect(gallery.open).toBe(false);
		gallery.trigger.click();
		expect(gallery.open).toBe(false);
		expect(pick).not.toHaveBeenCalled();
		gallery.state = undefined;
		expect(gallery.trigger.disabled).toBe(true);
	});

	it('cleans up on disconnect, isolates instances and styles each root once', () => {
		const first = mount();
		const second = mount();
		first.open = true;
		first.remove();
		expect(first.open).toBe(false);
		document.body.append(first);
		first.open = true;
		second.trigger.dispatchEvent(new Event('pointerdown', { bubbles: true }));
		expect(first.open).toBe(false);
		second.trigger.click();
		expect(second.open).toBe(true);
		expect(first.querySelectorAll('.gallery-view')).toHaveLength(1);
		const styles = document.querySelectorAll('style[data-office-gallery-styles]').length;
		const adopted = document.adoptedStyleSheets?.length ?? 0;
		expect(styles + adopted).toBe(1);
	});

	it('draws a dropdown trigger with its label and registered glyph', () => {
		const gallery = document.createElement('office-ui-gallery') as Gallery;
		gallery.setAttribute('icon', 'fill');
		gallery.state = STYLES;
		document.body.append(gallery);
		expect(gallery.trigger.textContent).toContain('Shape Styles');
		expect(gallery.trigger.querySelector('path')).not.toBeNull();
		expect(gallery.querySelector('.strip [data-gallery-item]')).toBeNull();
	});

	it('runs a command gallery as one button with its reason as the tooltip', () => {
		const command: OfficeGalleryState = {
			id: 'addShape',
			label: 'Add Shape',
			command: { icon: 'pencil', large: true },
			sections: [{ columns: 1, tileWidth: 1, tileHeight: 1, items: [{ id: 'run', label: 'Run' }] }],
		};
		const gallery = mount(command, false);
		const pick = vi.fn();
		gallery.addEventListener('office-gallery-pick', (e) => pick((e as CustomEvent).detail));
		expect(gallery.trigger.classList.contains('command')).toBe(true);
		expect(gallery.trigger.hasAttribute('aria-haspopup')).toBe(false);
		expect(gallery.trigger.textContent).toBe('Add Shape');
		expect(gallery.hasAttribute('data-command-large')).toBe(true);
		gallery.trigger.click();
		expect(pick).toHaveBeenCalledExactlyOnceWith({ gallery: 'addShape', itemId: 'run' });
		gallery.state = {
			...command,
			disabled: true,
			command: { ...command.command!, hint: 'Select a node' },
		};
		expect(gallery.trigger.disabled).toBe(true);
		expect(gallery.trigger.title).toBe('Select a node');
	});

	it('lets a product keep its event and hook names', () => {
		const Base = customElements.get('office-ui-gallery') as unknown as { new (): Gallery };
		class Product extends Base {
			static pickEvent = 'gallery-pick';
			static triggerAttribute = 'data-ribbon-gallery';
			static compactAttribute = 'data-compact';
		}
		customElements.define('product-gallery', Product as unknown as CustomElementConstructor);
		const gallery = document.createElement('product-gallery') as Gallery;
		gallery.setAttribute('mode', 'inline');
		gallery.state = STYLES;
		document.body.append(gallery);
		expect(gallery.trigger.getAttribute('data-ribbon-gallery')).toBe('shapeStyles');
		expect(gallery.querySelector('.tile')!.hasAttribute('data-compact')).toBe(true);
		const pick = vi.fn();
		gallery.addEventListener('gallery-pick', pick);
		gallery.querySelector<HTMLButtonElement>('.tile')!.click();
		expect(pick).toHaveBeenCalledOnce();
	});
});

describe('parseSvgPreview', () => {
	it('keeps SVG and strips script, handlers and javascript URLs', () => {
		const node = parseSvgPreview(
			document,
			'<svg xmlns="http://www.w3.org/2000/svg" onload="x()"><script>x()</script><a href="javascript:x()"><rect onclick="x()"/></a></svg>',
		)!;
		expect(node.localName).toBe('svg');
		expect(node.hasAttribute('onload')).toBe(false);
		expect(node.querySelector('script')).toBeNull();
		expect(node.querySelector('a')!.hasAttribute('href')).toBe(false);
		expect(node.querySelector('rect')!.hasAttribute('onclick')).toBe(false);
		expect(parseSvgPreview(document, '<div/>')).toBeNull();
	});
});
