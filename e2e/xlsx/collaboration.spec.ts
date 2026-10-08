import { expect, test, type Page } from '@playwright/test';
import {
	FRAMEWORKS,
	editor,
	goToCell,
	grid,
	newWorkbook,
	openSample,
	pageErrors,
	part,
	typeInActiveCell,
} from './helpers';

/**
 * Two windows of one browser share a workbook through File > Share. Without a server the editor
 * joins the room over a BroadcastChannel, which connects pages of one browser context, so each
 * pairing runs two pages built from different framework adapters in the same context.
 */
async function startSharing(page: Page, name: string, room: string) {
	await editor(page).evaluate((node, author) => {
		(node as unknown as { authorName: string }).authorName = author;
	}, name);
	await part(page, 'ribbon').locator('office-ui-ribbon-actions [part="share"]').click();
	const share = editor(page).locator('[data-backstage-page="share"]');
	await expect(share).toBeVisible();
	await share.locator('.xve-share-room').fill(room);
	await share.getByRole('button', { name: 'Start sharing', exact: true }).click();
	await expect(share.getByRole('button', { name: 'Stop sharing', exact: true })).toBeVisible();
	await part(page, 'backstage').locator('[data-backstage="back"]').click();
	await expect(share).toBeHidden();
}

/**
 * What the grid shows. The grid recycles hidden cell nodes (and hidden quadrants) that keep their
 * last text, so `toContainText` (textContent) would still see a cell that was cleared.
 */
const shown = (page: Page) => grid(page).innerText();

const sharing = (page: Page) =>
	editor(page).evaluate(
		(node) =>
			(node as unknown as { collaborationState: { people: { name: string }[] } }).collaborationState
				.people.length,
	);

for (const [index, host] of FRAMEWORKS.entries()) {
	const guest = FRAMEWORKS[(index + 1) % FRAMEWORKS.length]!;
	test(`${host} and ${guest} co-edit one workbook`, async ({ context }) => {
		const room = `e2e-${host}-${guest}-${Date.now().toString(36)}`;
		const ada = await context.newPage();
		const bob = await context.newPage();
		const errors = [pageErrors(ada), pageErrors(bob)];
		await openSample(ada, host);
		await startSharing(ada, 'Ada', room);
		await newWorkbook(bob, guest);
		await startSharing(bob, 'Bob', room);
		await expect.poll(() => sharing(ada)).toBe(2);
		await expect.poll(() => sharing(bob)).toBe(2);
		// The guest adopts the room's workbook (the sample's sheets).
		await expect(part(bob, 'sheet-tabs')).toContainText('Sales');

		await goToCell(ada, 'H2');
		await typeInActiveCell(ada, 'Ada was here');
		await expect.poll(() => shown(bob)).toContain('Ada was here');
		await goToCell(bob, 'H3');
		await typeInActiveCell(bob, 'Bob was here');
		await expect.poll(() => shown(ada)).toContain('Bob was here');

		// Ada's selection is outlined in Bob's grid with her name.
		await goToCell(ada, 'C4');
		await expect(grid(bob).locator('.xg-remote-tag', { hasText: 'Ada' }).first()).toBeVisible();
		await expect(part(bob, 'title-bar').locator('office-ui-presence')).toBeVisible();

		// Undo in one window reverts only that window's edit, everywhere.
		await grid(ada).focus();
		await ada.keyboard.press('Control+Z');
		await expect.poll(() => shown(bob)).not.toContain('Ada was here');
		await expect.poll(() => shown(ada)).not.toContain('Ada was here');
		await expect.poll(() => shown(ada)).toContain('Bob was here');
		await expect.poll(() => shown(bob)).toContain('Bob was here');

		// Stopping leaves the room: later edits stay local.
		await editor(bob).evaluate((node) =>
			(node as unknown as { stopCollaboration(): void }).stopCollaboration(),
		);
		await expect.poll(() => sharing(ada)).toBe(1);
		await goToCell(ada, 'H4');
		await typeInActiveCell(ada, 'after Bob left');
		await expect.poll(() => shown(ada)).toContain('after Bob left');
		await expect.poll(() => shown(bob)).not.toContain('after Bob left');
		expect(errors.flat()).toEqual([]);
	});
}
