import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';

// Run with Bun from e2e/pptx. The manifest and SVG files are generated local fixtures.
import { chromium } from '@playwright/test';

const manifest = resolve(process.argv[2]!);
const root = dirname(manifest);
const spec = JSON.parse(readFileSync(manifest, 'utf8')) as {
	cases: { id: string; width: number; height: number; sharedSvg: string }[];
};
const browser = await chromium.launch({ headless: true });
try {
	const captures: { id: string; sharedPng: string; svgSha256: string }[] = [];
	for (const item of spec.cases) {
		if (!/^[a-z0-9][a-z0-9-]*$/.test(item.id)) throw new Error('Invalid case id');
		if (![item.width, item.height].every((n) => Number.isInteger(n) && n > 0 && n <= 4096))
			throw new Error('Invalid capture dimensions');
		const page = await browser.newPage({
			viewport: { width: item.width, height: item.height },
			deviceScaleFactor: 1,
		});
		try {
			const svgSource = readFileSync(resolve(root, item.sharedSvg), 'utf8');
			// Never fetch remote fonts, images or scripts from a source fixture.
			await page.route(/^https?:/, (route) => route.abort());
			await page.setContent(
				`<style>html,body{margin:0;width:100%;height:100%;overflow:hidden}svg{display:block;width:100%;height:100%}</style>${svgSource}`,
			);
			await page.evaluate(async () => {
				await document.fonts.ready;
				await Promise.all(
					[...document.querySelectorAll('svg image')].map((node) => {
						const source = node.getAttribute('href') ?? node.getAttribute('xlink:href');
						if (!source) return Promise.resolve();
						const image = new Image();
						image.src = source;
						return image.decode();
					}),
				);
			});
			const sharedPng = `${item.id}.shared.png`;
			await page.screenshot({ path: resolve(root, sharedPng), omitBackground: true });
			captures.push({
				id: item.id,
				sharedPng,
				svgSha256: createHash('sha256').update(svgSource).digest('hex'),
			});
		} finally {
			await page.close();
		}
	}
	writeFileSync(
		resolve(root, 'shared.json'),
		JSON.stringify(
			{ browser: browser.version(), capture: 'Chromium SVG at 1x', cases: captures },
			null,
			2,
		),
	);
} finally {
	await browser.close();
}
