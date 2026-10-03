import type { Locator } from '@playwright/test';

/** Visio ribbon interactions shared by the browser specs. */
export async function taskPane(
	viewer: Locator,
	name: 'Pages' | 'Inspector' | 'Shape Data' | 'Layers' | 'Compatibility Notes',
): Promise<void> {
	await viewer.getByRole('tab', { name: 'View', exact: true }).click();
	await viewer.getByRole('button', { name: 'Task Panes', exact: true }).click();
	await viewer.locator(`office-ui-menu-item[label="${name}"]`).click();
}

export async function zoomPreset(viewer: Locator, percent: number): Promise<void> {
	await viewer.getByRole('tab', { name: 'View', exact: true }).click();
	await viewer.getByRole('button', { name: 'Zoom', exact: true }).click();
	await viewer.locator(`office-ui-menu-item[label="${percent}%"]`).click();
}

/** Home > Editing > Find > Find... opens the find bar, as Ctrl+F does. */
export async function openFind(viewer: Locator): Promise<Locator> {
	await viewer.getByRole('tab', { name: 'Home', exact: true }).click();
	const tools = viewer.locator('.ribbon-tools');
	if ((await tools.getAttribute('open')) === null) await tools.locator('summary').click();
	await viewer.getByRole('button', { name: 'Find', exact: true }).click();
	await viewer.locator('office-ui-menu-item[label="Find..."]').click();
	return viewer.getByRole('searchbox', { name: 'Search diagram text' });
}
