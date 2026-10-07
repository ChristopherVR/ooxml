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
