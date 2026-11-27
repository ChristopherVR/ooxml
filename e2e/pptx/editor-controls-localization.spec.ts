/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
import { expect, test } from '@playwright/test';

import { translationsEn } from '../../src/ui/src/pptx/i18n/translations-en';
import {
	translationsDe,
	translationsEs,
	translationsFr,
	translationsZhCN,
} from '../../viewers/pptx/packages/locales/src';
import { elementWithText, loadDeck, selectElement } from './support/deck';
import {
	openOptionsDialog,
	optionsAction,
	optionsCategory,
	pickOptionsEntry,
} from './support/settings-dialog';

// Keep the Drawing group expanded while switching the locale.
test.use({ viewport: { width: 3200, height: 1080 } });

for (const { dictionary, nativeLabel } of [
	{ dictionary: translationsDe, nativeLabel: 'Deutsch' },
	{ dictionary: translationsEs, nativeLabel: 'Español' },
	{ dictionary: translationsFr, nativeLabel: 'Français' },
	{ dictionary: translationsZhCN, nativeLabel: '简体中文' },
]) {
	test(`${nativeLabel} translates inspector navigation, alignment, and record commands`, async ({
		page,
	}) => {
		await loadDeck(page);
		const dialog = await openOptionsDialog(page, [
			translationsEn['pptx.options.title'],
			dictionary['pptx.options.title'],
		]);
		await optionsCategory(dialog, translationsEn['pptx.settings.language']).click();
		await pickOptionsEntry(dialog, nativeLabel);
		await optionsAction(dialog, dictionary['pptx.common.ok']).click();
		await expect(
			page.getByRole('tab', { name: dictionary['pptx.ribbon.tab.home'], exact: true }).first(),
		).toBeVisible();
		for (const key of [
			'pptx.documentProperties.statistics.elements',
			'pptx.inspector.properties',
		]) {
			await expect(
				page
					.getByRole('button', { name: dictionary[key], exact: true })
					.or(page.getByRole('tab', { name: dictionary[key], exact: true }))
					.first(),
			).toBeVisible();
		}
		await page.locator('[data-ribbon-control="home.drawing.arrange"] button').click();
		for (const key of [
			'pptx.ribbon.alignLeft',
			'pptx.ribbon.alignCenter',
			'pptx.ribbon.alignRight',
			'pptx.ribbon.alignTop',
			'pptx.ribbon.alignMiddle',
			'pptx.ribbon.alignBottom',
			'pptx.arrange.flipHorizontally',
			'pptx.ribbon.group',
		]) {
			await expect(
				page.getByRole('menuitem', { name: dictionary[key], exact: true }),
			).toBeVisible();
		}
		await page.keyboard.press('Escape');
		// The shared Home strips re-translate when the language changes at runtime.
		for (const key of [
			'pptx.arrange.copy',
			'pptx.home.newSlide',
			'pptx.drawing.shapes',
			'pptx.text.changeCase',
			'pptx.text.fontColor',
			'pptx.editing.find',
		]) {
			const name = dictionary[key] ?? translationsEn[key];
			await expect(page.getByRole('button', { name, exact: true }).first()).toBeVisible();
		}
		// The selects and menus rendered by the shared Home strips follow too.
		for (const key of [
			'pptx.ribbon.fontFamily',
			'pptx.ribbon.fontSize',
			'pptx.paragraph.columns',
		]) {
			const name = dictionary[key] ?? translationsEn[key];
			await expect(page.getByRole('combobox', { name, exact: true }).first()).toBeVisible();
		}
		const shape = elementWithText(page, 'Product Overview');
		await shape.click({ button: 'right' });
		await expect(
			page.locator('[data-item-id="duplicate"]').and(
				page.getByRole('menuitem', {
					name: dictionary['pptx.contextMenu.duplicate'],
					exact: true,
				}),
			),
		).toBeVisible();
		await page.keyboard.press('Escape');
		await selectElement(page, shape);
		await page.locator('[data-ribbon-contextual-tab="shapeFormat"]').click();
		await expect(
			page.getByRole('button', { name: dictionary['pptx.shape.mergeShapes'], exact: true }),
		).toBeVisible();
		await page
			.getByRole('tab', { name: dictionary['pptx.ribbon.tab.record'], exact: true })
			.click();
		for (const key of [
			'pptx.record.cameo',
			'pptx.slideShow.fromBeginning',
			'pptx.slideShow.fromCurrent',
			'pptx.record.clearTimings',
			'pptx.record.resetToCameo',
			'pptx.record.learnMore',
		]) {
			// The quick-access toolbar has icon-only commands with some of the same names.
			const command = page
				.getByRole('button', { name: dictionary[key], exact: true })
				.filter({ hasText: dictionary[key] });
			await expect(command).toBeVisible();
			await expect(command).toHaveText(dictionary[key]);
		}
		await expect(
			page
				.getByRole('button')
				.filter({ hasText: dictionary['pptx.customShows.addShow'] })
				.or(page.getByRole('button', { name: dictionary['pptx.ribbon.customShows'], exact: true }))
				.first(),
		).toBeVisible();
		if (nativeLabel === '简体中文') {
			await page.screenshot({ path: test.info().outputPath('record-zh-CN.png') });
		}
	});
}
