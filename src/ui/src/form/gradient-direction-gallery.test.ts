// @vitest-environment jsdom
import { expect, it } from 'vitest';
import { createGradientDirectionGallery } from './gradient-direction-gallery';

it('namespaces each preview, retains stop opacity and guards invalid or disabled picks', async () => {
	const first = createGradientDirectionGallery(document);
	const second = createGradientDirectionGallery(document);
	let picked = -1;
	const options = {
		gradient: {
			type: 'linear' as const,
			angle: 90,
			stops: [
				{ position: 0, color: '#FF0000', opacity: 0.63 },
				{ position: 100, color: '#FFFFFF' },
			],
		},
		disabled: false,
		label: 'Direction',
		translate: (label: string) => label,
		onPick: (angle: number) => {
			picked = angle;
		},
	};
	first.update(options);
	second.update(options);
	const items = first.element.state!.sections[0]!.items;
	expect(items).toHaveLength(8);
	expect(items.find((item) => item.applied)?.id).toBe('90');
	const previews = [...items, ...second.element.state!.sections[0]!.items].map((item) =>
		new DOMParser().parseFromString(item.preview!, 'image/svg+xml'),
	);
	const ids = previews.map((doc) => doc.querySelector('linearGradient')!.id);
	expect(new Set(ids).size).toBe(16);
	for (const preview of previews)
		expect(preview.querySelector('stop')!.getAttribute('stop-opacity')).toBe('0.63');
	const pick = (itemId: string) =>
		first.element.dispatchEvent(
			new CustomEvent('office-gallery-pick', { detail: { gallery: 'gradient-direction', itemId } }),
		);
	pick('315');
	expect(picked).toBe(315);
	pick('23');
	expect(picked).toBe(315);
	first.update({ ...options, disabled: true });
	pick('45');
	expect(picked).toBe(315);
	first.element.remove();
	second.element.remove();
});

it('previews native rectangular directions without flattening the gradient or accepting angle IDs', () => {
	const gallery = createGradientDirectionGallery(document);
	let picked = '';
	const options = {
		gradient: {
			type: 'radial' as const,
			path: 'rect',
			fillToRect: { l: 1, t: 0, r: 0, b: 1 },
			stops: [
				{ position: 0, color: '#FF0000', opacity: 0.63 },
				{ position: 100, color: '#FFFFFF' },
			],
		},
		disabled: false,
		label: 'Direction',
		translate: (label: string) => label,
		onPick: () => {
			throw new Error('Expected rectangular direction');
		},
		onRectPick: (direction: string) => {
			picked = direction;
		},
	};
	gallery.update(options);
	const items = gallery.element.state!.sections[0]!.items;
	expect(items).toHaveLength(5);
	expect(items.find((item) => item.applied)?.id).toBe('top-right');
	for (const item of items) {
		const preview = new DOMParser().parseFromString(item.preview!, 'image/svg+xml');
		expect(preview.querySelector('pattern')).not.toBeNull();
		expect(
			decodeURIComponent(preview.querySelector('image')!.getAttribute('href')!.split(',')[1]!),
		).toContain('<mask');
	}
	const pick = (itemId: string) =>
		gallery.element.dispatchEvent(
			new CustomEvent('office-gallery-pick', { detail: { gallery: 'gradient-direction', itemId } }),
		);
	pick('bottom-right');
	expect(picked).toBe('bottom-right');
	pick('90');
	expect(picked).toBe('bottom-right');
	gallery.update({ ...options, disabled: true });
	pick('center');
	expect(picked).toBe('bottom-right');
});

for (const path of ['circle', 'shape'])
	it(`reuses path directions and the shared ${path} painter`, () => {
		const gallery = createGradientDirectionGallery(document);
		let picked = '';
		gallery.update({
			gradient: {
				type: 'radial',
				path,
				stops: [
					{ position: 0, color: '#ff0000' },
					{ position: 100, color: '#ffffff' },
				],
				fillToRect: { l: 0.5, t: 0.5, r: 0.5, b: 0.5 },
			},
			disabled: false,
			label: 'Direction',
			translate: (label) => label,
			onPick: () => {
				throw new Error('Unexpected angle');
			},
			onPathPick: (direction) => {
				picked = direction;
			},
		});
		const items = gallery.element.state!.sections[0]!.items;
		expect(items).toHaveLength(5);
		expect(items.find((item) => item.applied)?.id).toBe('center');
		for (const item of items) {
			const doc = new DOMParser().parseFromString(item.preview!, 'image/svg+xml');
			expect(doc.querySelector(path === 'circle' ? 'radialGradient' : 'pattern')).not.toBeNull();
		}
		gallery.element.dispatchEvent(
			new CustomEvent('office-gallery-pick', { detail: { itemId: 'bottom-right' } }),
		);
		expect(picked).toBe('bottom-right');
	});
