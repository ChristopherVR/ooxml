import { expect, test, type Page } from '@playwright/test';
import { openSample } from './helpers';

// Host overrides of the shared theme reach the workbook where they are set: a container's
// `--office-accent` reaches the accent text inside it, and the documented `--xve-*` overrides win
// over the shared tokens.

/** Accent text (selected header, active sheet tab) and field boundaries, read inside the editor. */
async function paint(page: Page) {
	return page.evaluate(() => {
		const walk = function* (root: ParentNode): Generator<Element> {
			for (const el of root.querySelectorAll('*')) {
				yield el;
				if (el.shadowRoot) yield* walk(el.shadowRoot);
			}
		};
		const all = [...walk(document.querySelector('xlsx-editor')!.shadowRoot!)];
		const read = (selector: string, property: string) => {
			const el = all.find((candidate) => candidate.matches(selector));
			return el ? getComputedStyle(el).getPropertyValue(property) : 'missing';
		};
		return {
			header: read('.xg-hd-sel', 'color'),
			sheet: read('.xst-tab.xst-active', 'color'),
			combo: read('.ribbon-combo', 'border-bottom-color'),
			formula: read('.xfb-input', 'border-bottom-color'),
		};
	});
}

async function set(page: Page, target: 'container' | 'host', name: string, value: string) {
	await page.evaluate(
		([target, name, value]) => {
			const editor = document.querySelector<HTMLElement>('xlsx-editor')!;
			// The container is the innermost theme scope around the editor's chrome: the editor itself
			// when it carries its own `data-office-theme`, which resets the palette, else its parent.
			const container = editor.hasAttribute('data-office-theme') ? editor : editor.parentElement!;
			(target === 'host' ? editor : container).style.setProperty(name!, value!);
		},
		[target, name, value],
	);
}

test('a container accent and the --xve-* overrides reach accent text and field boundaries', async ({
	page,
}) => {
	await openSample(page);

	await set(page, 'container', '--office-accent', 'rgb(170, 0, 0)');
	await expect
		.poll(async () => {
			const { header, sheet } = await paint(page);
			return { header, sheet };
		})
		.toEqual({ header: 'rgb(170, 0, 0)', sheet: 'rgb(170, 0, 0)' });
	await set(page, 'container', '--office-control-border', 'rgb(0, 0, 255)');
	await expect
		.poll(async () => {
			const { combo, formula } = await paint(page);
			return { combo, formula };
		})
		.toEqual({ combo: 'rgb(0, 0, 255)', formula: 'rgb(0, 0, 255)' });

	await set(page, 'host', '--xve-primary', 'rgb(0, 100, 0)');
	await expect.poll(async () => (await paint(page)).header).toBe('rgb(0, 100, 0)');
	// The selection colour also colours the selected header and the active sheet's name.
	await set(page, 'host', '--xve-selection-border', 'rgb(0, 0, 160)');
	await expect
		.poll(async () => {
			const { header, sheet } = await paint(page);
			return { header, sheet };
		})
		.toEqual({ header: 'rgb(0, 0, 160)', sheet: 'rgb(0, 0, 160)' });
	await set(page, 'host', '--xve-input', 'rgb(120, 0, 120)');
	await expect
		.poll(async () => {
			const { combo, formula } = await paint(page);
			return { combo, formula };
		})
		.toEqual({ combo: 'rgb(120, 0, 120)', formula: 'rgb(120, 0, 120)' });
});
