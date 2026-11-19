import { expect, test, type Page } from '@playwright/test';
import { openSample } from './helpers';

// Host overrides of the shared theme reach the editor where they are set: a container's
// `--office-accent` reaches the accent text inside it, and the documented `--dve-*` overrides win
// over the shared tokens.

/** The selected ribbon tab's text and the font name box's boundary, read inside the editor. */
async function paint(page: Page) {
	return page.evaluate(() => {
		const walk = function* (root: ParentNode): Generator<Element> {
			for (const el of root.querySelectorAll('*')) {
				yield el;
				if (el.shadowRoot) yield* walk(el.shadowRoot);
			}
		};
		const all = [...walk(document.querySelector('docx-editor')!.shadowRoot!)];
		const tab = all.find((el) => el.matches('[role="tab"][aria-selected="true"]'));
		const combo = all.find((el) => el.matches('.ribbon-combo.font-family-combo'));
		return {
			tab: tab ? getComputedStyle(tab).color : 'missing',
			combo: combo ? getComputedStyle(combo).borderBottomColor : 'missing',
		};
	});
}

async function set(page: Page, target: 'container' | 'host', name: string, value: string) {
	await page.evaluate(
		([target, name, value]) => {
			const editor = document.querySelector<HTMLElement>('docx-editor')!;
			// The container is the innermost theme scope around the editor's chrome: the editor itself
			// when it carries its own `data-office-theme`, which resets the palette, else its parent.
			const container = editor.hasAttribute('data-office-theme') ? editor : editor.parentElement!;
			(target === 'host' ? editor : container).style.setProperty(name!, value!);
		},
		[target, name, value],
	);
}

test('a container accent and the --dve-* overrides reach accent text and field boundaries', async ({
	page,
}) => {
	await openSample(page);

	await set(page, 'container', '--office-accent', 'rgb(170, 0, 0)');
	await expect.poll(async () => (await paint(page)).tab).toBe('rgb(170, 0, 0)');
	await set(page, 'container', '--office-control-border', 'rgb(0, 0, 255)');
	await expect.poll(async () => (await paint(page)).combo).toBe('rgb(0, 0, 255)');

	await set(page, 'host', '--dve-primary', 'rgb(0, 100, 0)');
	await expect.poll(async () => (await paint(page)).tab).toBe('rgb(0, 100, 0)');
	await set(page, 'host', '--dve-primary-ink', 'rgb(0, 0, 160)');
	await expect.poll(async () => (await paint(page)).tab).toBe('rgb(0, 0, 160)');
	await set(page, 'host', '--dve-input', 'rgb(120, 0, 120)');
	await expect.poll(async () => (await paint(page)).combo).toBe('rgb(120, 0, 120)');
});
