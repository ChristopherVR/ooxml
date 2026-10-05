import type { Locator } from '@playwright/test';

/** Visio ribbon interactions shared by the browser specs. */
export async function taskPane(
	viewer: Locator,
	name: 'Shapes' | 'Inspector' | 'Shape Data' | 'Layers' | 'Compatibility Notes' | 'Pan & Zoom',
): Promise<void> {
	await viewer.getByRole('tab', { name: 'View', exact: true }).click();
	await viewer.getByRole('button', { name: 'Task Panes', exact: true }).click();
	await viewer
		.locator(`office-ui-menu-button[data-menu="task-panes"] office-ui-menu-item[label="${name}"]`)
		.click();
}

export async function zoomPreset(viewer: Locator, percent: number): Promise<void> {
	await viewer.getByRole('tab', { name: 'View', exact: true }).click();
	await viewer.getByRole('button', { name: 'Zoom', exact: true }).click();
	await viewer
		.locator(`office-ui-menu-button[data-menu="zoom"] office-ui-menu-item[label="${percent}%"]`)
		.click();
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

/** Visio's File backstage. */
export async function fileBackstage(
	viewer: Locator,
	item:
		| 'info'
		| 'new'
		| 'open'
		| 'save-as'
		| 'print'
		| 'share'
		| 'export'
		| 'account'
		| 'options' = 'info',
): Promise<void> {
	const backstage = viewer.locator('.backstage');
	if (!(await backstage.isVisible())) await viewer.locator('office-ui-ribbon .file').click();
	await viewer.locator(`[data-backstage-item="${item}"]`).click();
}

/** File > Save: downloads the VSDX copy (disabled for model-only and legacy drawings). */
export function saveCommand(viewer: Locator): Locator {
	return viewer.locator('[data-backstage-item="save"]');
}

/** File > Save As > Download a copy. */
export async function downloadCopy(viewer: Locator): Promise<Locator> {
	await fileBackstage(viewer, 'save-as');
	return viewer.locator('[data-backstage-page="save-as"] [data-backstage-action="download"]');
}

/** File > Export > Export the current page as SVG. */
export async function exportSvgCommand(viewer: Locator): Promise<Locator> {
	await fileBackstage(viewer, 'export');
	return viewer.locator('[data-backstage-page="export"] [data-backstage-action="export-svg"]');
}

/** File > New > the demo's slotted Sample workflow template. */
export async function loadSampleTemplate(viewer: Locator): Promise<void> {
	await fileBackstage(viewer, 'new');
	await viewer.locator('.template-card').click();
}
