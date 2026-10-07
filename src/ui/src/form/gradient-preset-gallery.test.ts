// @vitest-environment jsdom
import { expect, it } from 'vitest';
import { officeGradientPresetFill } from 'ooxml-core/diagram';
import { createGradientPresetGallery } from './gradient-preset-gallery';

it('reuses native paint, namespaces thumbnails and guards invalid and disabled picks', () => {
	const first = createGradientPresetGallery(document),
		second = createGradientPresetGallery(document);
	let picked = -1;
	const options = {
		fill: officeGradientPresetFill(5),
		disabled: false,
		label: 'Preset gradients',
		translate: (label: string) => label,
		onPick: (id: number) => {
			picked = id;
		},
	};
	first.update(options);
	second.update(options);
	const items = first.element.state!.sections[0]!.items;
	expect(items).toHaveLength(24);
	expect(items.find((item) => item.applied)?.label).toBe('Horizon');
	const previews = [...items, ...second.element.state!.sections[0]!.items].map((item) =>
		new DOMParser().parseFromString(item.preview!, 'image/svg+xml'),
	);
	expect(new Set(previews.map((doc) => doc.querySelector('linearGradient')!.id)).size).toBe(48);
	expect(previews[4]!.querySelectorAll('stop')).toHaveLength(10);
	const pick = (id: string) =>
		first.element.dispatchEvent(
			new CustomEvent('office-gallery-pick', {
				detail: { gallery: 'gradient-presets', itemId: id },
			}),
		);
	pick('16');
	expect(picked).toBe(16);
	pick('025');
	expect(picked).toBe(16);
	first.update({ ...options, disabled: true });
	pick('5');
	expect(picked).toBe(16);
});
