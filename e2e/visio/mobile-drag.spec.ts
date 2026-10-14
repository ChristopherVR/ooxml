import { expect, test } from '@playwright/test';
import type { VisioViewerElement } from 'ooxml-ui/visio';
import { openDemo } from './demo-page';

test.use({ viewport: { width: 390, height: 844 }, hasTouch: true, isMobile: true });
for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	test(`${framework}: native touch moves a Visio shape through implicit capture transfer and supports undo`, async ({
		page,
	}) => {
		await openDemo(page, `${framework === 'vanilla' ? '/demo/' : `/demo-${framework}/`}?sample=1`);
		const viewer = page.locator('visio-viewer');
		const shape = viewer.locator('svg.paper [data-shape-id]').first();
		await expect(shape).toBeVisible();
		const pin = () =>
			viewer.evaluate((node) => {
				const s = (node as VisioViewerElement).controller.state.document!.pages[0]!.shapes[0]!;
				return { x: s.rotation!.pinX, y: s.rotation!.pinY };
			});
		const before = await pin();
		const box = (await shape.boundingBox())!;
		const x = Math.round(box.x + box.width / 2),
			y = Math.round(box.y + box.height / 2);
		const cdp = await page.context().newCDPSession(page);
		await cdp.send('Input.dispatchTouchEvent', { type: 'touchStart', touchPoints: [{ x, y }] });
		for (let i = 1; i <= 10; i++) {
			await cdp.send('Input.dispatchTouchEvent', {
				type: 'touchMove',
				touchPoints: [{ x: x + i * 5, y: y + i * 3 }],
			});
			await page.waitForTimeout(25);
		}
		await cdp.send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [] });
		await expect.poll(pin).not.toEqual(before);
		await expect
			.poll(() =>
				viewer.evaluate((node) => (node as VisioViewerElement).controller.state.edit.busy),
			)
			.toBe(false);
		await viewer.evaluate((node) => (node as VisioViewerElement).controller.undo());
		await expect.poll(pin).toEqual(before);
		await shape.tap();
		const handle = viewer.locator('[data-resize-handle="se"]');
		await expect(handle).toBeVisible();
		const resizeBefore = await shape.boundingBox();
		const handleBox = (await handle.boundingBox())!;
		const hx = Math.round(handleBox.x + handleBox.width / 2),
			hy = Math.round(handleBox.y + handleBox.height / 2);
		await cdp.send('Input.dispatchTouchEvent', {
			type: 'touchStart',
			touchPoints: [{ x: hx, y: hy }],
		});
		for (let i = 1; i <= 10; i++) {
			await cdp.send('Input.dispatchTouchEvent', {
				type: 'touchMove',
				touchPoints: [{ x: hx + i * 3, y: hy + i * 2 }],
			});
			await page.waitForTimeout(25);
		}
		await cdp.send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [] });
		await expect
			.poll(async () => (await shape.boundingBox())?.width ?? 0)
			.toBeGreaterThan(resizeBefore!.width + 20);
		await expect
			.poll(() =>
				viewer.evaluate((node) => (node as VisioViewerElement).controller.state.edit.busy),
			)
			.toBe(false);
		await viewer.evaluate((node) => (node as VisioViewerElement).controller.undo());
		await expect.poll(pin).toEqual(before);
		await expect
			.poll(async () => (await shape.boundingBox())?.width ?? 0)
			.toBeCloseTo(resizeBefore!.width, 1);
		await expect
			.poll(() =>
				viewer.evaluate((node) => (node as VisioViewerElement).controller.state.edit.busy),
			)
			.toBe(false);
		// Empty canvas remains pannable after selection mode takes ownership of touch.
		await viewer.evaluate((node) => {
			(node as VisioViewerElement).zoom = 2;
		});
		const viewport = viewer.locator('.viewport');
		await expect
			.poll(() => viewport.evaluate((node) => node.scrollHeight > node.clientHeight))
			.toBe(true);
		await viewport.evaluate(async (node) => {
			node.scrollTop = 100;
			node.scrollLeft = 100;
			// Native touch uses compositor hit testing, which must catch up with the scroll.
			await new Promise<void>((resolve) =>
				requestAnimationFrame(() => requestAnimationFrame(() => resolve())),
			);
		});
		const port = (await viewport.boundingBox())!;
		const px = Math.round(port.x + 30),
			py = Math.round(port.y + 60);
		// Verify the gesture starts on empty canvas, so a shape drag cannot mask a pan.
		await expect
			.poll(() =>
				viewer.evaluate(
					(node, point) => {
						const target = node.shadowRoot!.elementFromPoint(point.x, point.y);
						return (
							!!target?.closest('.viewport') &&
							!target.closest('[data-shape-id], [data-resize-handle], [data-rotation-handle]')
						);
					},
					{ x: px, y: py },
				),
			)
			.toBe(true);
		const scrollBefore = await viewport.evaluate((node) => node.scrollTop);
		await cdp.send('Input.dispatchTouchEvent', {
			type: 'touchStart',
			touchPoints: [{ x: px, y: py }],
		});
		for (let i = 1; i <= 10; i++) {
			await cdp.send('Input.dispatchTouchEvent', {
				type: 'touchMove',
				touchPoints: [{ x: px, y: py - i * 3 }],
			});
			await page.waitForTimeout(25);
		}
		await cdp.send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [] });
		await expect
			.poll(() => viewport.evaluate((node) => node.scrollTop))
			.toBeGreaterThan(scrollBefore + 20);
		await expect.poll(pin).toEqual(before);
	});
