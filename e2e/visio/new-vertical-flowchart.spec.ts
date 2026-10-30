import { expect, test } from '@playwright/test';
import type { VisioViewerElement } from 'ooxml-ui/visio';
import { openDemo } from './demo-page';

test('File New offers a vertical Cross-Functional Flowchart with shapes in their lanes', async ({
	page,
}) => {
	await page.setViewportSize({ width: 1600, height: 1000 });
	await openDemo(page, '/demo/?sample=1');
	const viewer = page.locator('visio-viewer');
	await viewer.locator('office-ui-ribbon .file').click();
	await viewer.locator('[data-backstage-item="new"]').click();
	await viewer
		.locator('[data-backstage-action="new-template:cross-functional-flowchart-vertical"]')
		.click();
	await expect
		.poll(() => viewer.evaluate((node) => (node as VisioViewerElement).fileName))
		.toBe('New drawing.vsdx');
	const lanes = await viewer.evaluate((node) => {
		const shapes = (node as VisioViewerElement).controller.state.document!.pages[0]!.shapes;
		const frame = shapes.find((shape) => shape.id === '50')!.structure as {
			swimlane: { options: { direction: string } };
		};
		const lane = (id: string) => shapes.find((shape) => shape.id === id)!;
		return {
			direction: frame.swimlane.options.direction,
			columns: lane('10').width < lane('10').height,
			members: ['10', '11', '12'].map((id) => {
				const structure = lane(id).structure;
				return structure?.type === 'container' ? structure.memberIds.length : 0;
			}),
		};
	});
	expect(lanes).toEqual({ direction: 'vertical', columns: true, members: [1, 2, 1] });
});
