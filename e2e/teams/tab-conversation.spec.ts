import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import type { TeamsApp } from '../../src/ui/src/teams/app/teams-app';

test('shares one tab conversation with Posts and preserves workbook edits while discussing', async ({
	context,
}) => {
	test.setTimeout(120_000);
	const ada = await context.newPage(),
		bob = await context.newPage();
	const bytes = await readFile(
		new URL('../../src/core/xlsx/__fixtures__/openpyxl-features.xlsx', import.meta.url),
	);
	await context.route('https://files.test/Budget.xlsx', (route) =>
		route.fulfill({
			body: bytes,
			contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
		}),
	);
	const room = `tab-chat-${Date.now()}`;
	await ada.goto(`/?local=1&name=Ada&room=${room}`);
	await bob.goto(`/?local=1&name=Bob&room=${room}`);
	await bob.locator('teams-app').evaluate((element) => ((element as TeamsApp).userId = 'bob'));
	for (const page of [ada, bob])
		await expect(page.getByRole('heading', { name: '# General', exact: true })).toBeVisible({
			timeout: 15_000,
		});
	await ada.locator('teams-app').evaluate(async (element) => {
		const client = (element as TeamsApp).client!;
		client.workspace.chat.post(client.getState().selectedChannelId, {
			text: 'Shared budget',
			attachments: [{ name: 'Budget.xlsx', kind: 'xlsx', url: 'https://files.test/Budget.xlsx' }],
		});
	});
	await ada.getByRole('button', { name: 'Add tab', exact: true }).click();
	await ada.getByRole('button', { name: 'Choose Excel', exact: true }).click();
	await ada.getByRole('radio', { name: /^Budget.xlsx/ }).check();
	await expect(
		ada.getByRole('checkbox', { name: 'Post to the channel about this tab', exact: true }),
	).toBeChecked();
	await ada.getByRole('button', { name: 'Save', exact: true }).click();
	await expect(ada.getByRole('button', { name: 'Edit workbook', exact: true })).toBeVisible({
		timeout: 60_000,
	});
	await ada.getByRole('button', { name: 'Edit workbook', exact: true }).click();
	const formula = ada.getByRole('textbox', { name: 'Formula Bar', exact: true });
	await formula.click();
	await formula.press('ControlOrMeta+A');
	await formula.pressSequentially('Keep this budget edit');
	await formula.press('Enter');
	const nameBox = ada.getByRole('combobox', { name: 'Name Box', exact: true });
	await nameBox.fill('A1');
	await nameBox.press('Enter');
	await expect(formula).toHaveValue('Keep this budget edit');
	await ada.getByRole('button', { name: 'Show tab conversation', exact: true }).click();
	await expect(ada.getByRole('heading', { name: 'Thread', exact: true })).toBeFocused();
	await expect(formula).toHaveValue('Keep this budget edit');
	const adaThread = ada.getByRole('complementary', { name: 'Thread', exact: true });
	await adaThread
		.getByRole('textbox', { name: 'Message', exact: true })
		.fill('Budget needs approval');
	await adaThread.getByRole('textbox', { name: 'Message', exact: true }).press('Enter');
	await expect(bob.getByText('Discuss the Budget.xlsx tab here.', { exact: true })).toBeVisible();
	await bob.getByRole('button', { name: 'Open tab', exact: true }).click();
	await bob.getByRole('button', { name: 'Show tab conversation', exact: true }).click();
	const bobThread = bob.getByRole('complementary', { name: 'Thread', exact: true });
	await expect(bobThread.getByText('Budget needs approval', { exact: true })).toBeVisible();
	await bobThread.getByRole('textbox', { name: 'Message', exact: true }).fill('Approved by Bob');
	await bobThread.getByRole('textbox', { name: 'Message', exact: true }).press('Enter');
	await expect(adaThread.getByText('Approved by Bob', { exact: true })).toBeVisible();
	await ada.screenshot({ path: test.info().outputPath('xlsx-tab-conversation.png') });
	await adaThread.getByRole('textbox', { name: 'Message', exact: true }).fill('Keep reply draft');
	await adaThread.getByRole('button', { name: 'Close thread', exact: true }).click();
	await expect(
		ada.getByRole('button', { name: 'Show tab conversation', exact: true }),
	).toBeFocused();
	await expect(formula).toHaveValue('Keep this budget edit');
	await ada.getByRole('button', { name: 'Show tab conversation', exact: true }).click();
	await expect(adaThread.getByRole('textbox', { name: 'Message', exact: true })).toHaveValue(
		'Keep reply draft',
	);
	await bob.getByRole('tab', { name: 'Posts', exact: true }).click();
	await bob.getByRole('button', { name: '2 replies', exact: true }).click();
	await expect(bobThread.getByText('Approved by Bob', { exact: true })).toBeVisible();
	await bob.getByRole('button', { name: 'Close thread', exact: true }).click();
	await bob.getByRole('button', { name: 'Add tab', exact: true }).click();
	await bob.getByRole('button', { name: 'Choose Website', exact: true }).click();
	await bob.getByRole('textbox', { name: 'Tab name', exact: true }).fill('Quiet site');
	await bob
		.getByRole('textbox', { name: 'Tab website URL', exact: true })
		.fill('https://example.com');
	await bob
		.getByRole('checkbox', { name: 'Post to the channel about this tab', exact: true })
		.uncheck();
	await bob.getByRole('button', { name: 'Save', exact: true }).click();
	await expect(bob.getByRole('tab', { name: 'Quiet site', exact: true })).toHaveAttribute(
		'aria-selected',
		'true',
	);
	await bob.getByRole('tab', { name: 'Posts', exact: true }).click();
	await expect(
		bob.getByText('Discuss the Quiet site tab here.', { exact: true }),
	).not.toBeVisible();
	await bob.getByRole('tab', { name: 'Quiet site', exact: true }).click();
	await bob.getByRole('button', { name: 'Show tab conversation', exact: true }).click();
	await expect(
		bobThread.getByText('Discuss the Quiet site tab here.', { exact: true }),
	).toBeVisible();
	await bob.setViewportSize({ width: 390, height: 844 });
	await expect(bobThread).toBeVisible();
	await bob.screenshot({ path: test.info().outputPath('tab-conversation-mobile.png') });
	await bob.getByRole('button', { name: 'Close thread', exact: true }).click();
	await bob.getByRole('tab', { name: 'Posts', exact: true }).click();
	await expect(bob.getByText('Discuss the Quiet site tab here.', { exact: true })).toBeVisible();
	await context.close();
});
