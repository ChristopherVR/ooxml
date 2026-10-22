/**
 * Touch gestures whose timing does not depend on how fast the page handles them.
 *
 * `touchscreen.tap()` sends the next touch only once the page has handled the previous one, and
 * the browser stamps each touch when it is sent. On a loaded runner the page can take 300ms to
 * handle the tap that selects an element, so two taps 80ms apart in a spec reach the page more
 * than half a second apart and miss every binding's double-tap window, which a person tapping
 * twice on a slow phone never does: the device stamps the touches when they happen.
 *
 * @module e2e/support/touch
 */
import type { Page } from '@playwright/test';

/** How long each finger stays down, and the gap between the two taps, as a person taps. */
const PRESS_MS = 30;
const GAP_MS = 120;

/**
 * Double-tap at a viewport point with touches stamped `GAP_MS` apart, the same CDP touch events
 * `touchscreen.tap()` sends (so the browser still derives its pointer, tap and click events
 * from them) but with explicit timestamps.
 */
export async function doubleTap(page: Page, x: number, y: number): Promise<void> {
	const session = await page.context().newCDPSession(page);
	try {
		const start = Date.now() / 1000;
		const point = [{ x, y }];
		const at = (ms: number) => start + ms / 1000;
		for (const tap of [0, PRESS_MS + GAP_MS]) {
			await session.send('Input.dispatchTouchEvent', {
				type: 'touchStart',
				touchPoints: point,
				timestamp: at(tap),
			});
			await session.send('Input.dispatchTouchEvent', {
				type: 'touchEnd',
				touchPoints: [],
				timestamp: at(tap + PRESS_MS),
			});
		}
	} finally {
		await session.detach();
	}
}
