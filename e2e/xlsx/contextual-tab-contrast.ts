import { expect, type Locator } from '@playwright/test';
import { contrastRatio } from '../../src/core/color/contrast';

export async function expectContextualTabContrast(tab: Locator): Promise<void> {
	await expect(tab).toHaveAttribute('data-contextual', '');
	const measurements = await tab.evaluate((node) => {
		const ribbon = node.getRootNode();
		if (!(ribbon instanceof ShadowRoot)) throw new Error('Expected a ribbon shadow root');
		const editor = ribbon.host.getRootNode();
		if (!(editor instanceof ShadowRoot)) throw new Error('Expected an editor shadow root');
		const host = editor.host;
		const original = host.getAttribute('data-office-theme');
		const canvas = new OffscreenCanvas(1, 1).getContext('2d')!;
		const hex = (color: string) => {
			canvas.clearRect(0, 0, 1, 1);
			canvas.fillStyle = color;
			canvas.fillRect(0, 0, 1, 1);
			return [...canvas.getImageData(0, 0, 1, 1).data]
				.slice(0, 3)
				.map((channel) => channel.toString(16).padStart(2, '0'))
				.join('');
		};
		try {
			return ['light', 'dark'].map((theme) => {
				host.setAttribute('data-office-theme', theme);
				const foreground = getComputedStyle(node).color;
				const background = getComputedStyle(ribbon.host).backgroundColor;
				return { theme, foreground: hex(foreground), background: hex(background) };
			});
		} finally {
			if (original === null) host.removeAttribute('data-office-theme');
			else host.setAttribute('data-office-theme', original);
		}
	});
	for (const { theme, foreground, background } of measurements) {
		expect(
			contrastRatio(foreground, background),
			`${theme}: #${foreground} on #${background}`,
		).toBeGreaterThanOrEqual(4.5);
	}
}
