/**
 * axe on the main surfaces in light, dark and on a phone: a critical or serious violation beyond
 * the checked-in baseline fails. Each run records what it found (`axe.json` in its output folder)
 * so `node check-baseline.mjs --update` can lower the baseline after a fix.
 */
import { writeFileSync } from 'node:fs';
import { axeViolations, overBaseline, slack } from './support/axe';
import { openSurface, SURFACES, test, expect, type Layout, type Mode } from './support/suite';

const MODES: Record<string, { mode: Mode; layout: Layout; mobile?: true }> = {
	light: { mode: 'light', layout: 'office' },
	dark: { mode: 'dark', layout: 'office' },
	mobile: { mode: 'light', layout: 'simple', mobile: true },
};

for (const [name, { mode, layout, mobile }] of Object.entries(MODES)) {
	test.describe(`axe, ${name}`, () => {
		test.use({
			mode,
			layout,
			...(mobile
				? {
						viewport: { width: 390, height: 844 },
						hasTouch: true,
						isMobile: true,
						deviceScaleFactor: 2,
					}
				: { viewport: { width: 1440, height: 900 } }),
		});
		for (const surface of SURFACES) {
			const key = `${surface}/${name}`;
			test(`${surface} has no critical or serious violations beyond the baseline`, async ({
				page,
			}, info) => {
				await openSurface(page, surface);
				const violations = await axeViolations(page);
				writeFileSync(info.outputPath('axe.json'), JSON.stringify({ key, violations }, null, '\t'));
				const unused = slack(key, violations);
				if (unused.length)
					info.annotations.push({
						type: 'baseline can shrink',
						description: `${key}: ${unused.join(', ')} (node check-baseline.mjs --update)`,
					});
				const over = overBaseline(key, violations).map(
					(v) =>
						`${v.rule} (${v.impact}): ${v.nodes} nodes, e.g. ${v.targets.join(' | ')}\n  ${v.help}`,
				);
				expect(over, `new axe violations on ${key}`).toEqual([]);
			});
		}
	});
}
