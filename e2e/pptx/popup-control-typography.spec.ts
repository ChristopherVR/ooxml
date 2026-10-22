/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
/**
 * The shared popup controls keep the typography their own stylesheet gives them.
 *
 * The viewer stylesheet's reset is scoped to `[data-pptx-viewer]`,
 * `[data-pptx-viewport]` and `[data-pptx-scope]`, and an outermost scope starts
 * from the page's base text (16px, line-height 1.5). A popup a binding mounts on
 * `<body>` is its own outermost scope. If the marker sat on the shadow-DOM control
 * itself, those page rules would beat the control's `:host` typography and the
 * context menu would grow from 12px to 16px. The expected values are read from
 * the control's own `:host` rules rather than hard-coded, so a token change in the
 * control does not need this spec to change.
 *
 * Run: bunx playwright test popup-control-typography
 */
import { expect, test } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';

import { stageElements } from './support/context-menu';
import { loadDeck, slideStage } from './support/deck';
import { openOptionsDialog } from './support/settings-dialog';

test.use({ viewport: { width: 1440, height: 1100 } });

interface Typography {
	/** Whether the control's own `:host` rules set its font size or line height. */
	declared: boolean;
	actual: { fontSize: string; lineHeight: string };
	expected: { fontSize: string; lineHeight: string };
}

/**
 * The control's computed font size and line height, beside what its own stylesheet makes them.
 *
 * The control's `:host` rules (those whose argument the host matches) are replayed as the
 * inline style of a probe placed next to the host, so the probe inherits what the host inherits
 * and resolves the same tokens, with nothing from the page able to override it. A control whose
 * `:host` rules set no typography should inherit its parent's, so that is the expectation then.
 */
async function controlTypography(control: Locator): Promise<Typography> {
	return control.evaluate((host): Typography => {
		const root = host.shadowRoot;
		const parent = host.parentElement;
		if (!root || !parent) {
			throw new Error(`${host.localName} has no shadow root or no parent`);
		}
		const splitTopLevel = (selector: string): string[] => {
			const parts: string[] = [];
			let depth = 0;
			let start = 0;
			for (let index = 0; index < selector.length; index += 1) {
				const char = selector[index];
				depth += char === '(' ? 1 : char === ')' ? -1 : 0;
				if (char === ',' && depth === 0) {
					parts.push(selector.slice(start, index).trim());
					start = index + 1;
				}
			}
			parts.push(selector.slice(start).trim());
			return parts;
		};
		// 0 for `:host`, 1 for `:host(<selector>)`, which is more specific; -1 when not the host.
		const hostRank = (selector: string): number => {
			const ranks = splitTopLevel(selector).map((part) => {
				if (part === ':host') {
					return 0;
				}
				const argument = /^:host\((.+)\)$/u.exec(part)?.[1];
				return argument && host.matches(argument) ? 1 : -1;
			});
			return Math.max(...ranks);
		};
		const hostRules: { rank: number; style: CSSStyleDeclaration }[] = [];
		const collect = (rules: CSSRuleList): void => {
			for (const rule of Array.from(rules)) {
				if (rule instanceof CSSMediaRule) {
					if (matchMedia(rule.media.mediaText).matches) {
						collect(rule.cssRules);
					}
				} else if (rule instanceof CSSStyleRule) {
					const rank = hostRank(rule.selectorText);
					if (rank >= 0) {
						hostRules.push({ rank, style: rule.style });
					}
				}
			}
		};
		for (const sheet of [...root.adoptedStyleSheets, ...Array.from(root.styleSheets)]) {
			collect(sheet.cssRules);
		}
		// Stable sort: source order within a rank, as the cascade does.
		hostRules.sort((a, b) => a.rank - b.rank);

		const probe = document.createElement('div');
		const hostStyle = getComputedStyle(host);
		let declared = false;
		for (const { style } of hostRules) {
			for (const name of ['font', 'font-family', 'font-size', 'line-height']) {
				const value = style.getPropertyValue(name);
				if (!value) {
					continue;
				}
				// Tokens resolve to what the host sees, wherever they are set (its own rules or a theme).
				for (const [, token] of value.matchAll(/var\(\s*(--[\w-]+)/gu)) {
					probe.style.setProperty(token, hostStyle.getPropertyValue(token));
				}
				probe.style.setProperty(name, value);
				declared ||= name !== 'font-family';
			}
		}
		parent.insertBefore(probe, host);
		const read = (element: Element) => {
			const style = getComputedStyle(element);
			return { fontSize: style.fontSize, lineHeight: style.lineHeight };
		};
		const expected = read(declared ? probe : parent);
		probe.remove();
		return { declared, actual: read(host), expected };
	});
}

async function openDeck(page: Page): Promise<void> {
	await loadDeck(page);
	await slideStage(page).waitFor();
	await page.waitForTimeout(500);
}

test('the slide context menu keeps the font size and line height its own styles set', async ({
	page,
}) => {
	await openDeck(page);
	const shape = stageElements(page).filter({ hasText: 'Product Overview' }).first();
	const box = await shape.boundingBox();
	expect(box).not.toBeNull();
	await page.mouse.click(box!.x + box!.width / 2, box!.y + box!.height / 2, { button: 'right' });
	const menu = page.locator('pptx-ui-context-menu');
	await expect(menu).toHaveCount(1);
	await expect(menu.getByRole('menu')).toBeVisible();
	const typography = await controlTypography(menu);
	// The menu sets its own text (`font` on its controlled `:host`), so this is not vacuous.
	expect(typography.declared).toBeTruthy();
	expect(typography.actual).toStrictEqual(typography.expected);
});

test('the options dialog keeps the font size and line height its own styles set', async ({
	page,
}) => {
	await openDeck(page);
	await openOptionsDialog(page, ['Options']);
	const dialog = page.locator('pptx-ui-options-dialog');
	await expect(dialog).toHaveCount(1);
	const typography = await controlTypography(dialog);
	expect(typography.actual).toStrictEqual(typography.expected);
});
