import { test, expect, type Page } from '@playwright/test';
import { sampleSettled } from './demo-page';

// A shared drawing merges per shape property: two windows that change DIFFERENT properties of the
// same shape at the same time must both keep their change. The element API drives the edits
// because a drag and an inline text edit through the UI would not be concurrent.
interface Probe {
	pageId: string;
	shapeId: string;
	text: string;
	transform: string;
}

async function openPair(context: import('@playwright/test').BrowserContext, room: string) {
	const a = await context.newPage();
	const b = await context.newPage();
	await a.goto(`/demo/?embed=1&sample=1&share=${room}`);
	await sampleSettled(a);
	await expect(a.locator('visio-viewer svg.paper')).toContainText('Start with an idea');
	// The second window joins late and adopts what the room holds.
	await b.goto(`/demo/?embed=1&share=${room}`);
	await expect(b.locator('visio-viewer svg.paper')).toContainText('Start with an idea');
	return { a, b };
}

/** The shape `wanted` (default: the first with text) of the first page; null while it is absent. */
function probe(page: Page, wanted?: string): Promise<Probe | null> {
	return page.locator('visio-viewer').evaluate((element: any, id: string | undefined) => {
		const drawing = element.document.pages[0];
		const shape = id
			? drawing.shapes.find((s: { id: string }) => s.id === id)
			: drawing.shapes.find((s: { text?: { plainText?: string } }) => s.text?.plainText);
		if (!shape) return null;
		return {
			pageId: drawing.id,
			shapeId: shape.id,
			text: shape.text?.plainText ?? '',
			transform: JSON.stringify(shape.transform),
		};
	}, wanted);
}

test('concurrent edits to different properties of one shape both survive', async ({ context }) => {
	const { a, b } = await openPair(context, `e2e-merge-${Date.now().toString(36)}`);
	const before = (await probe(a))!;
	await expect
		.poll(() => probe(b, before.shapeId).then((shape) => shape?.transform), { timeout: 30_000 })
		.toBe(before.transform);

	// A moves the shape while B changes its text, in the same instant.
	await Promise.all([
		a
			.locator('visio-viewer')
			.evaluate(
				(element: any, target: Probe) =>
					element.applyEdits([
						{ type: 'move-shape', pageId: target.pageId, shapeId: target.shapeId, x: 6, y: 5 },
					]),
				before,
			),
		b
			.locator('visio-viewer')
			.evaluate(
				(element: any, target: Probe) =>
					element.replacePlainText(target.pageId, target.shapeId, 'Edited in the other window'),
				before,
			),
	]);

	for (const page of [a, b]) {
		await expect(page.locator('visio-viewer svg.paper')).toContainText(
			'Edited in the other window',
			{
				timeout: 30_000,
			},
		);
		await expect
			.poll(async () => (await probe(page, before.shapeId))?.transform, { timeout: 30_000 })
			.not.toBe(before.transform);
	}
	const onA = (await probe(a, before.shapeId))!;
	const onB = (await probe(b, before.shapeId))!;
	expect(onA.transform).toBe(onB.transform);
	expect(onA.text).toContain('Edited in the other window');
	expect(onB.text).toContain('Edited in the other window');
});

test('a collaborator selection shows as an outline and a cursor in the other window', async ({
	context,
}) => {
	const { a, b } = await openPair(context, `e2e-presence-${Date.now().toString(36)}`);
	const target = (await probe(a))!;

	await a
		.locator('visio-viewer')
		.evaluate(
			(element: any, shape: Probe) =>
				element.selectShapes([{ id: shape.shapeId, name: '', pageId: shape.pageId }]),
			target,
		);
	const outline = b.locator(
		`visio-viewer svg.paper [data-presence-overlay] [data-presence-selection="${target.shapeId}"]`,
	);
	await expect(outline).toHaveCount(1, { timeout: 30_000 });
	// The outline carries the sender's colour.
	await expect(outline.locator('rect')).toHaveAttribute('stroke', /^#[\da-f]{3,8}$/i);

	// Moving the pointer over A's canvas shows A's cursor in B.
	const canvas = a.locator('visio-viewer svg.paper');
	const box = await canvas.boundingBox();
	await a.mouse.move(box!.x + box!.width / 2, box!.y + box!.height / 2, { steps: 4 });
	await expect(b.locator('visio-viewer svg.paper [data-presence-cursor]')).toHaveCount(1, {
		timeout: 30_000,
	});

	// Clearing the selection removes the outline.
	await a.locator('visio-viewer').evaluate((element: any) => element.clearSelection());
	await expect(outline).toHaveCount(0, { timeout: 30_000 });
});
