/**
 * Reflow (WCAG 1.4.10) at 320x256 CSS pixels and 360x225 (1280x900 at 400% zoom): the page never
 * scrolls sideways, every control can still be reached, and the document keeps a usable share of
 * the window. A grid or canvas may scroll in two dimensions inside itself; the page may not.
 */
import { gate } from './support/known';
import { expect, openSurface, PRODUCTS, test, type Layout, type Surface } from './support/suite';

const SIZES = { '320x256': { width: 320, height: 256 }, '360x225': { width: 360, height: 225 } };
/** The document must keep at least this share of the window's height and width. */
const MIN_SHARE = { height: 1 / 4, width: 1 / 2 };

const DOCUMENT: Partial<Record<Surface, string>> = {
	docx: '[aria-label="Document page"]',
	xlsx: '[role="grid"]',
	pptx: '[data-pptx-viewport]',
	vsdx: '[aria-label="Diagram canvas"]',
};

interface Reflow {
	pageScrollsSideways: boolean;
	unreachable: string[];
	document?: { width: number; height: number };
}

/** Measures the page: sideways scroll, controls no scrolling can bring into view, document box. */
const measure = (page: import('@playwright/test').Page, documentSelector: string | undefined) =>
	page.evaluate((selector): Reflow => {
		const all = globalThis.__a11y.queryAll('*');
		const root = document.documentElement;
		const up = (node: Element): Element | null =>
			node.assignedSlot ?? node.parentElement ?? (node.getRootNode() as ShadowRoot).host ?? null;
		const scrolls = (element: Element) => {
			for (let node = up(element); node && node !== root; node = up(node)) {
				const style = getComputedStyle(node);
				if (
					(/(auto|scroll)/u.test(style.overflowX) && node.scrollWidth > node.clientWidth + 1) ||
					(/(auto|scroll)/u.test(style.overflowY) && node.scrollHeight > node.clientHeight + 1)
				)
					return true;
			}
			return false;
		};
		const focusable = all.filter(
			(element) =>
				element.matches(
					'a[href], button:not([disabled]), input:not([disabled]):not([type="hidden"]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])',
				) &&
				element.getClientRects().length > 0 &&
				getComputedStyle(element).visibility !== 'hidden' &&
				!element.closest('[inert], [aria-hidden="true"]'),
		);
		const pageScrolls = root.scrollWidth > innerWidth + 1 || root.scrollHeight > innerHeight + 1;
		const unreachable = focusable
			.filter((element) => {
				const box = element.getBoundingClientRect();
				const inView =
					box.right > 0 && box.bottom > 0 && box.left < innerWidth && box.top < innerHeight;
				return !inView && !scrolls(element) && !pageScrolls;
			})
			.map((element) =>
				(element.getAttribute('aria-label') || element.textContent || element.localName)
					.trim()
					.slice(0, 30),
			);
		// What the user sees of the document: its box clipped by the window and by every ancestor
		// that clips its content.
		const target = selector ? all.find((element) => element.matches(selector)) : undefined;
		let visible: { left: number; top: number; right: number; bottom: number } | undefined;
		if (target) {
			const box = target.getBoundingClientRect();
			visible = {
				left: Math.max(box.left, 0),
				top: Math.max(box.top, 0),
				right: Math.min(box.right, innerWidth),
				bottom: Math.min(box.bottom, innerHeight),
			};
			for (let node = up(target); node && node !== root; node = up(node)) {
				const style = getComputedStyle(node);
				if (style.overflowX === 'visible' && style.overflowY === 'visible') continue;
				const clip = node.getBoundingClientRect();
				visible = {
					left: Math.max(visible.left, clip.left),
					top: Math.max(visible.top, clip.top),
					right: Math.min(visible.right, clip.right),
					bottom: Math.min(visible.bottom, clip.bottom),
				};
			}
		}
		return {
			pageScrollsSideways: root.scrollWidth > innerWidth + 1,
			unreachable,
			...(visible && {
				document: {
					width: Math.max(0, visible.right - visible.left),
					height: Math.max(0, visible.bottom - visible.top),
				},
			}),
		};
	}, documentSelector);

const CASES: [Surface, Layout][] = [
	...PRODUCTS.flatMap((product): [Surface, Layout][] => [
		[product, 'office'],
		[product, 'simple'],
	]),
	['teams', 'office'],
	['chat', 'office'],
];

for (const [size, viewport] of Object.entries(SIZES))
	test.describe(`reflow at ${size}`, () => {
		test.use({ viewport });
		for (const [surface, layout] of CASES)
			test.describe(`${surface}${DOCUMENT[surface] ? `, ${layout} layout` : ''}`, () => {
				test.use({ layout });
				const key = `reflow/${size}/${surface}${DOCUMENT[surface] ? `/${layout}` : ''}`;

				test('the page does not scroll sideways and every control can be reached', async ({
					page,
				}) => {
					await gate(`${key}/controls`, async () => {
						await openSurface(page, surface);
						const reflow = await measure(page, DOCUMENT[surface]);
						expect(reflow.pageScrollsSideways, 'the page scrolls sideways').toBe(false);
						expect(
							reflow.unreachable,
							'controls outside the window that nothing scrolls to',
						).toEqual([]);
					});
				});

				const selector = DOCUMENT[surface];
				if (selector)
					test('the document keeps a usable share of the window', async ({ page }) => {
						await gate(`${key}/document`, async () => {
							await openSurface(page, surface);
							const { document } = await measure(page, selector);
							expect(document, `the document (${selector}) is on the page`).toBeDefined();
							expect(
								document?.height ?? 0,
								'the document is at least a quarter of the window tall',
							).toBeGreaterThanOrEqual(viewport.height * MIN_SHARE.height);
							expect(
								document?.width ?? 0,
								'the document is at least half the window wide',
							).toBeGreaterThanOrEqual(viewport.width * MIN_SHARE.width);
						});
					});
			});
	});
