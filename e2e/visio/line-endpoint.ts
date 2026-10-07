import { expect, type Page } from '@playwright/test';

/** Real pointer interaction in drawing inches, shared by native geometry/paint tests. */
export async function dragLineEndpoint(
	page: Page,
	shapeId: string,
	endpoint: 'begin' | 'end',
	x: number,
	y: number,
) {
	const viewer = page.locator('visio-viewer');
	const line = viewer.locator(`[data-shape-id="${shapeId}"]`);
	await line.focus();
	await line.press('Enter');
	const handle = viewer.locator(
		`[data-line-shape-id="${shapeId}"][data-line-endpoint="${endpoint}"]`,
	);
	await expect(handle).toBeVisible();
	const before = (await line.getAttribute('transform'))!;
	const origin = (await handle.boundingBox())!;
	const target = await viewer.evaluate(
		(node, coordinates) => {
			const element = node as unknown as {
				document: import('ooxml-core/visio').VisioDocument;
				shadowRoot: ShadowRoot;
			};
			const page = element.document.pages[0]!;
			const svg = element.shadowRoot.querySelector<SVGSVGElement>('svg.paper')!;
			const point = new DOMPoint(
				coordinates.x * (page.drawingToPageScale ?? 1),
				page.height - coordinates.y * (page.drawingToPageScale ?? 1),
			).matrixTransform(svg.getScreenCTM()!);
			return { x: point.x, y: point.y };
		},
		{ x, y },
	);
	await page.mouse.move(origin.x + origin.width / 2, origin.y + origin.height / 2);
	await page.mouse.down();
	await page.mouse.move(target.x, target.y, { steps: 4 });
	await expect(viewer.locator('.endpoint-preview')).toBeVisible();
	await page.mouse.up();
	await expect(viewer.locator('.endpoint-preview')).toHaveCount(0);
	await expect(line).not.toHaveAttribute('transform', before);
	const after = (await line.getAttribute('transform'))!;
	return { line, handle, before, after };
}
