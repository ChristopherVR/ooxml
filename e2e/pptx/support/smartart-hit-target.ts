import type { Locator } from '@playwright/test';

export async function unobstructedNodePoint(node: Locator): Promise<{ x: number; y: number }> {
	const point = await node.evaluate((element) => {
		const box = element.getBoundingClientRect();
		for (const yFraction of [0.8, 0.65, 0.5, 0.35, 0.2]) {
			for (const xFraction of [0.25, 0.5, 0.75, 0.1, 0.9]) {
				const x = box.left + box.width * xFraction;
				const y = box.top + box.height * yFraction;
				if (document.elementFromPoint(x, y)?.closest('[data-smartart-node-id]') === element) {
					return { x, y };
				}
			}
		}
		return null;
	});
	if (!point) {
		throw new Error('SmartArt node has no unobstructed browser hit target');
	}
	return point;
}
