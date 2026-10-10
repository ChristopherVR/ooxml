import { expect, test, type Locator } from '@playwright/test';
import JSZip from 'jszip';
import type { VisioViewerElement } from 'ooxml-ui/visio';
import { createVsdxFixture } from './fixture.mjs';
import { downloadCopy } from './ribbon';
import { openDemo } from './demo-page';

async function fixture(scale: number, protectedSecond = false) {
	const zip = await JSZip.loadAsync(await createVsdxFixture('Paint source'));
	const cell = (name: string, value: number | string) => `<Cell N="${name}" V="${value}"/>`;
	const shape = (id: number) =>
		`<Shape ID="${id}" NameU="Paint ${id}">${cell('Width', 2 / scale)}${cell('Height', 1 / scale)}${cell('PinX', (id * 2 + 1) / scale)}${cell('PinY', 7 / scale)}${cell('LocPinX', 1 / scale)}${cell('LocPinY', 0.5 / scale)}${cell('FillPattern', 1)}${cell('FillForegnd', '#DAEFE6')}${cell('FillBkgnd', '#FFFFFF')}${cell('FillForegndTrans', 0)}${cell('FillBkgndTrans', 0)}${cell('LinePattern', id)}${cell('LineColor', '#000000')}${cell('LineColorTrans', 0)}${cell('LineWeight', 1 / 72 / scale)}${id === 2 && protectedSecond ? cell('LockFormat', 1) : ''}<Section N="Geometry" IX="0"><Row T="MoveTo" IX="1">${cell('X', 0)}${cell('Y', 0)}</Row><Row T="LineTo" IX="2">${cell('X', 2 / scale)}${cell('Y', 0)}</Row><Row T="LineTo" IX="3">${cell('X', 2 / scale)}${cell('Y', 1 / scale)}</Row><Row T="LineTo" IX="4">${cell('X', 0)}${cell('Y', 1 / scale)}</Row><Row T="LineTo" IX="5">${cell('X', 0)}${cell('Y', 0)}</Row></Section><Text>Paint ${id}</Text></Shape>`;
	const page = await zip.file('visio/pages/page1.xml')!.async('string');
	zip.file(
		'visio/pages/page1.xml',
		page.replace(/<Shapes>.*?<\/Shapes>/su, `<Shapes>${shape(1)}${shape(2)}</Shapes>`),
	);
	const pages = await zip.file('visio/pages/pages.xml')!.async('string');
	zip.file(
		'visio/pages/pages.xml',
		pages
			.replace('N="PageWidth" V="8.5"', `N="PageWidth" V="${8.5 / scale}"`)
			.replace('N="PageHeight" V="11"', `N="PageHeight" V="${11 / scale}"`)
			.replace('</PageSheet>', `${cell('DrawingScale', 1)}${cell('PageScale', scale)}</PageSheet>`),
	);
	return zip.generateAsync({ type: 'nodebuffer' });
}
const idle = (viewer: Locator) =>
	expect
		.poll(() =>
			viewer.evaluate((node) => {
				const state = (node as VisioViewerElement).controller.state;
				return state.loading || state.edit.busy;
			}),
		)
		.toBe(false);
async function inventory(viewer: Locator) {
	await idle(viewer);
	return viewer.evaluate((node) => {
		const controller = (node as VisioViewerElement).controller;
		return {
			bytes: Array.from(controller.exportVsdx().bytes),
			selected: controller.state.selectedShapes.map((shape) => shape.id),
			shapes: controller.state.document!.pages[0]!.shapes.map((shape) => ({
				id: shape.id,
				pattern: shape.style.fillPatternIndex,
				background: shape.style.fillBackgroundColor,
				foregroundOpacity: shape.style.fillForegroundOpacity,
				backgroundOpacity: shape.style.fillBackgroundOpacity,
				line: shape.style.linePattern,
				lineOpacity: shape.style.lineColorOpacity,
				color: shape.style.lineColor,
			})),
		};
	});
}
async function selectAll(viewer: Locator) {
	await viewer.getByRole('tab', { name: 'Home', exact: true }).click();
	await viewer.getByRole('button', { name: 'Select', exact: true }).click();
	await viewer.locator('[command="select-all"]').click();
}
async function lineAction(viewer: Locator, id: string, nested = false) {
	await viewer.getByRole('button', { name: 'Line', exact: true }).click();
	if (nested)
		await viewer
			.locator('[data-menu="line-pattern"]')
			.getByRole('button', { name: 'Pattern', exact: true })
			.click();
	await viewer.locator(`[command="${id}"]`).click();
	await idle(viewer);
}
async function properties(viewer: Locator) {
	await viewer.getByRole('button', { name: 'Fill', exact: true }).click();
	await viewer.locator('[command="fill-options"]').click();
	await expect(
		viewer.locator('.paint-properties-dialog:not(.format-shape-dialog)'),
	).toHaveAttribute('open', '');
}
async function choose(viewer: Locator, field: string, label: string) {
	const control = viewer.locator(`[data-paint-field="${field}"]`);
	await control.getByRole('combobox').click();
	await control.getByRole('option', { name: label, exact: true }).click();
}

for (const [index, framework] of [
	'vanilla',
	'react',
	'vue',
	'angular',
	'svelte',
	'solid',
].entries()) {
	test(`${framework}: paint menus and properties save atomic selection edits with exact history`, async ({
		page,
	}) => {
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await page.setViewportSize({ width: 1600, height: 1000 });
		await openDemo(
			page,
			framework === 'vanilla' ? '/demo/?sample=1' : `/demo-${framework}/?sample=1`,
		);
		const viewer = page.locator('visio-viewer'),
			scale = [0.5, 1, 2][index % 3]!;
		const load = async (buffer: Buffer) => {
			await page.locator('#file').setInputFiles({
				name: 'paint.vsdx',
				mimeType: 'application/vnd.ms-visio.drawing',
				buffer,
			});
			await expect(page.locator('#file-name')).toHaveText('paint.vsdx');
			await idle(viewer);
		};
		await load(await fixture(scale));
		await selectAll(viewer);
		expect((await inventory(viewer)).selected).toEqual(['1', '2']);
		await lineAction(viewer, 'line-pattern-0');
		await expect(viewer.locator('svg.paper [data-shape-id="1"] > path')).toHaveAttribute(
			'stroke',
			'none',
		);
		await lineAction(viewer, 'line-blue');
		expect((await inventory(viewer)).shapes.map((shape) => shape.line)).toEqual([0, 0]);
		const before = await inventory(viewer);
		await properties(viewer);
		await choose(viewer, 'fillPattern', 'Pattern 24');
		await choose(viewer, 'linePattern', 'Pattern 23');
		await viewer
			.getByRole('textbox', { name: 'Fill background color', exact: true })
			.fill('#123456');
		await viewer
			.getByRole('spinbutton', { name: 'Fill transparency (%)', exact: true })
			.fill('17.5');
		await viewer
			.getByRole('spinbutton', { name: 'Line transparency (%)', exact: true })
			.fill('62.5');
		expect((await inventory(viewer)).bytes).toEqual(before.bytes);
		await viewer.locator('[command="paint-apply"] button').click();
		await idle(viewer);
		await expect(
			viewer.locator('.paint-properties-dialog:not(.format-shape-dialog)'),
		).not.toHaveAttribute('open', '');
		const accepted = await inventory(viewer);
		expect(accepted.selected).toEqual(['1', '2']);
		for (const shape of accepted.shapes) {
			expect(shape.pattern).toBe(24);
			expect(shape.background).toBe('#123456');
			expect(shape.line).toBe(23);
			expect(shape.foregroundOpacity).toBeCloseTo(0.825, 12);
			expect(shape.backgroundOpacity).toBeCloseTo(0.825, 12);
			expect(shape.lineOpacity).toBeCloseTo(0.375, 12);
		}
		const path = viewer.locator('svg.paper [data-shape-id="1"] > path');
		await expect(path).toHaveAttribute('stroke-opacity', '0.375');
		await expect(path).toHaveAttribute('stroke-dasharray', /\S+/);
		await expect(path).toHaveAttribute('fill', /^url\(#/u);
		await viewer.locator('.qat [data-command="undo"]').click();
		await idle(viewer);
		expect((await inventory(viewer)).bytes).toEqual(before.bytes);
		await viewer.locator('.qat [data-command="redo"]').click();
		await idle(viewer);
		expect((await inventory(viewer)).bytes).toEqual(accepted.bytes);
		await properties(viewer);
		await viewer.getByRole('spinbutton', { name: 'Line transparency (%)', exact: true }).fill('90');
		await page.keyboard.press('Escape');
		expect((await inventory(viewer)).bytes).toEqual(accepted.bytes);
		const downloadCommand = await downloadCopy(viewer),
			pending = page.waitForEvent('download');
		await downloadCommand.click();
		const download = await pending;
		await page.locator('#file').setInputFiles((await download.path())!);
		await idle(viewer);
		expect((await inventory(viewer)).shapes).toEqual(accepted.shapes);
		await viewer.locator('[data-shape-id="1"]').click({ button: 'right' });
		await viewer.locator('[command="ctx-format"]').click();
		await expect(
			viewer.locator('.paint-properties-dialog:not(.format-shape-dialog)'),
		).toHaveAttribute('open', '');
		await viewer.locator('[command="paint-cancel"] button').click();
		await load(await fixture(scale, true));
		await selectAll(viewer);
		const protectedSource = await inventory(viewer);
		await properties(viewer);
		await viewer.getByRole('spinbutton', { name: 'Line transparency (%)', exact: true }).fill('25');
		await viewer.locator('[command="paint-apply"] button').click();
		await idle(viewer);
		await expect(viewer.locator('[data-paint-error]')).toContainText('protection is active');
		expect((await inventory(viewer)).bytes).toEqual(protectedSource.bytes);
		expect((await inventory(viewer)).selected).toEqual(['1', '2']);
		await viewer.locator('[command="paint-cancel"] button').click();
		expect(errors).toEqual([]);
	});
}
