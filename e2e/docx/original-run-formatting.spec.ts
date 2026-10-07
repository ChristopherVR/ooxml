import { fileURLToPath } from 'node:url';
import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import { fileInput } from './helpers';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';

const fixture = (name: string, suffix: string) =>
	fileURLToPath(
		new URL(
			`../../src/core/docx/__fixtures__/review-formatting/${name}-${suffix}.docx`,
			import.meta.url,
		),
	);
for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	for (const name of ['bold', 'multiple'])
		test(`${framework}: Original renders native prior ${name} run formatting without resolving revisions`, async ({
			page,
		}) => {
			await page.setViewportSize({ width: 2400, height: 1000 });
			const errors: string[] = [];
			page.on('pageerror', (error) => errors.push(error.message));
			await page.goto(`/?framework=${framework}`);
			const input = await fileInput(page);
			const editor = page.locator('docx-editor');
			const body = editor.locator('.ProseMirror');
			const appearance = () =>
				body.evaluate((element) => {
					const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT);
					const result = [];
					for (let text = walker.nextNode(); text; text = walker.nextNode()) {
						const parent = text.parentElement!;
						const style = getComputedStyle(parent);
						const decorations = new Set<string>();
						let background = 'transparent';
						for (
							let ancestor: HTMLElement | null = parent;
							ancestor && ancestor !== element;
							ancestor = ancestor.parentElement
						) {
							const computed = getComputedStyle(ancestor);
							for (const line of computed.textDecorationLine.split(' '))
								if (line !== 'none') decorations.add(line);
							if (background === 'transparent' && computed.backgroundColor !== 'rgba(0, 0, 0, 0)')
								background = computed.backgroundColor;
						}
						for (const character of text.textContent ?? '')
							result.push({
								character,
								font: style.fontFamily,
								size: style.fontSize,
								bold: style.fontWeight,
								italic: style.fontStyle,
								color: style.color,
								decoration: [...decorations].sort(),
								background,
								transform: style.textTransform,
							});
					}
					return result;
				});
			await input.setInputFiles(fixture(name, 'before'));
			await expect(body).toContainText('Format me');
			await page.evaluate(() => document.fonts.ready);
			const before = await appearance();
			await input.setInputFiles(fixture(name, 'tracked'));
			await expect(body.locator('.dve-revision-format-markup')).not.toHaveCount(0);
			const current = await appearance();
			expect(current).not.toEqual(before);
			const source = await editor.evaluate(
				(element) => (element as DocxEditorElement).documentModel,
			);
			await editor.getByRole('tab', { name: 'Review', exact: true }).click();
			const mode = editor.getByRole('combobox', { name: 'Display for review', exact: true });
			await mode.selectOption('original', { force: true });
			await expect.poll(appearance).toEqual(before);
			expect(
				await editor.evaluate((element) => (element as DocxEditorElement).documentModel),
			).toEqual(source);
			const bytes = await editor.evaluate(async (element) =>
				Array.from(await (element as DocxEditorElement).saveBytes()),
			);
			const xml = await (
				await JSZip.loadAsync(new Uint8Array(bytes))
			)
				.file('word/document.xml')!
				.async('string');
			expect(xml).toContain('rPrChange');
			await mode.selectOption('final', { force: true });
			await expect.poll(appearance).toEqual(current);
			expect(errors).toEqual([]);
		});
