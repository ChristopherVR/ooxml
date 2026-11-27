import { expect, test, type Page } from '@playwright/test';
import { createWorkbook, putCell, saveXlsx } from 'ooxml-core/xlsx';
import type { XlsxEditorElement } from '../../viewers/xlsx/packages/web-component/src/index';
import { RICH_METADATA_EXT, RICH_NS, RICH_PARTS } from '../../src/core/xlsx/rich-values';
import { miniPackage, R, ws, X } from '../../src/core/xlsx/write/mini-package-fixtures';
import {
	FRAMEWORKS,
	editor,
	goToCell,
	grid,
	landingFileInput,
	newWorkbook,
	openLanding,
	ribbon,
	typeInActiveCell,
} from './helpers';

const PNG = Buffer.from(
	'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGNoTgkFAAKqAT2gS4eLAAAAAElFTkSuQmCC',
	'base64',
);

async function openPictures(page: Page): Promise<void> {
	await ribbon(page).getByRole('tab', { name: 'Insert', exact: true }).click();
	const pictures = editor(page).getByRole('button', { name: 'Pictures', exact: true });
	if (!(await pictures.isVisible()))
		await ribbon(page).getByRole('button', { name: 'Illustrations', exact: true }).click();
	await pictures.click();
}

test('inserts and converts a cell picture and blocks IMAGE without a host fetcher', async ({
	page,
}) => {
	await openLanding(page);
	await newWorkbook(page);
	await openPictures(page);
	const chooser = page.waitForEvent('filechooser');
	await editor(page).getByText('Place in Cell', { exact: true }).click();
	await (await chooser).setFiles({ name: 'pixel.png', mimeType: 'image/png', buffer: PNG });
	const picture = grid(page).locator('.xg-c:not([hidden]) > img');
	await expect(picture).toBeVisible();
	await expect(picture).toHaveCSS('object-fit', 'contain');
	await expect
		.poll(() =>
			picture.evaluate((image) => image instanceof HTMLImageElement && image.naturalWidth),
		)
		.toBe(1);
	await grid(page).focus();
	await grid(page).press('Shift+F10');
	await editor(page).getByText('Alt Text', { exact: true }).last().click();
	await editor(page).getByRole('textbox', { name: 'Alt Text' }).fill('Small square');
	await editor(page).getByRole('button', { name: 'OK', exact: true }).click();
	await expect(picture).toHaveAttribute('alt', 'Small square');
	await grid(page).focus();
	await grid(page).press('Shift+F10');
	await editor(page).getByText('Place over Cells', { exact: true }).click();
	await expect(picture).toHaveCount(0);
	await expect(grid(page).getByRole('img', { name: 'Small square' })).toBeVisible();
	await openPictures(page);
	await editor(page).getByText('Convert to In-Cell Picture', { exact: true }).click();
	await expect(picture).toHaveAttribute('alt', 'Small square');
	await goToCell(page, 'B1');
	await typeInActiveCell(page, '=IMAGE("https://example.test/picture.png")');
	await expect(
		grid(page)
			.locator('.xg-tx')
			.filter({ hasText: /^#BLOCKED!$/ }),
	).toBeVisible();
});

for (const framework of FRAMEWORKS) {
	test(`${framework} uses the host IMAGE provider on an open workbook`, async ({ page }) => {
		await openLanding(page, framework);
		await newWorkbook(page);
		await editor(page).evaluate(
			(element, bytes) => {
				const xlsx = element as XlsxEditorElement;
				xlsx.imageProvider = {
					resolve: () => true,
					getImage: () => ({ bytes: Uint8Array.from(bytes) }),
				};
			},
			[...PNG],
		);
		await typeInActiveCell(page, '=IMAGE("https://example.test/picture.png","Host image")');
		const picture = grid(page).locator('.xg-c:not([hidden]) > img[alt="Host image"]');
		await expect(picture).toBeVisible();
		await expect(grid(page).getByRole('gridcell', { name: 'Host image' })).toBeVisible();
		await editor(page).evaluate((element) => {
			const xlsx = element as XlsxEditorElement;
			xlsx.imageProvider = undefined;
			xlsx.calculateNow();
		});
		await expect(picture).toHaveCount(0);
		await expect(
			grid(page).getByRole('gridcell', { name: '#BLOCKED!', exact: true }),
		).toBeVisible();
	});
}

test('uses the host IMAGE provider when a file requests calculation on open', async ({ page }) => {
	await newWorkbook(page);
	const workbook = createWorkbook();
	workbook.fullCalcOnLoad = true;
	putCell(workbook.sheets[0]!, 0, 0, {
		value: { error: '#VALUE!' },
		formula: 'IMAGE("https://example.test/picture.png","On-open image")',
	});
	await editor(page).evaluate(
		async (element, { file, png }) => {
			const xlsx = element as XlsxEditorElement;
			let ready = false;
			xlsx.imageProvider = {
				resolve: () => {
					if (ready) return true;
					queueMicrotask(() => {
						ready = true;
						xlsx.calculateNow();
					});
					return { error: '#BUSY!' };
				},
				getImage: () => ({ bytes: Uint8Array.from(png) }),
			};
			await xlsx.load(Uint8Array.from(file), 'image-on-open.xlsx');
		},
		{ file: [...(await saveXlsx(workbook))], png: [...PNG] },
	);
	await expect(grid(page).locator('.xg-c:not([hidden]) > img[alt="On-open image"]')).toBeVisible();
});

test('keeps Place in Cell available while sharing', async ({ page }) => {
	await openLanding(page);
	await newWorkbook(page);
	await editor(page).evaluate(
		(element, roomId) =>
			(element as XlsxEditorElement).startCollaboration({ roomId, user: { name: 'Ada' } }),
		`pictures-${Date.now()}`,
	);
	await openPictures(page);
	// In-cell pictures are shared through the room (up to 4 MB each).
	const command = editor(page).getByRole('menuitem', { name: 'Place in Cell', exact: true });
	await expect(command).toBeEnabled();
	await expect(command).not.toHaveAttribute('aria-description');
	await page.keyboard.press('Escape');
	await editor(page).evaluate((element) => (element as XlsxEditorElement).stopCollaboration());
	await openPictures(page);
	await expect(command).toBeEnabled();
});

test('shows a cached data type icon and opens its card from the keyboard', async ({ page }) => {
	const { data, structures } = RICH_PARTS;
	const bytes = await miniPackage({
		sheets: [
			{
				name: 'Cities',
				xml: ws(
					'<row r="1"><c r="A1" vm="1" t="e"><v>#VALUE!</v></c></row>' +
						'<row r="2"><c r="A2" vm="2"><v>42</v></c></row>' +
						'<row r="3"><c r="A3" vm="3" t="e"><v>#VALUE!</v></c></row>',
					'',
					'<conditionalFormatting sqref="A2"><cfRule type="iconSet" priority="1"><iconSet iconSet="3Arrows"><cfvo type="num" val="0"/><cfvo type="num" val="10"/><cfvo type="num" val="20"/></iconSet></cfRule></conditionalFormatting>',
				),
			},
		],
		workbookRels: `<Relationship Id="meta" Type="${R}/sheetMetadata" Target="metadata.xml"/><Relationship Id="data" Type="${data.rel}" Target="richData/${data.file}"/><Relationship Id="structures" Type="${structures.rel}" Target="richData/${structures.file}"/>`,
		parts: {
			'xl/metadata.xml': `<metadata xmlns="${X}" xmlns:r="${RICH_NS}"><metadataTypes count="1"><metadataType name="XLRICHVALUE" minSupportedVersion="120000"/></metadataTypes><futureMetadata name="XLRICHVALUE" count="3">${[0, 1, 2].map((i) => `<bk><extLst><ext uri="${RICH_METADATA_EXT}"><r:rvb i="${i}"/></ext></extLst></bk>`).join('')}</futureMetadata><valueMetadata count="3">${[0, 1, 2].map((i) => `<bk><rc t="1" v="${i}"/></bk>`).join('')}</valueMetadata></metadata>`,
			[`xl/richData/${data.file}`]: `<rvData xmlns="${RICH_NS}" count="3"><rv s="0"><v>Seattle</v><v>Washington</v><v>780000</v></rv><rv s="1"><v>Example Inc</v><v>EXM</v><v>42</v></rv><rv s="2"><v>Generic record</v></rv></rvData>`,
			[`xl/richData/${structures.file}`]: `<rvStructures xmlns="${RICH_NS}" count="3"><s t="_entity"><k n="_DisplayString" t="s"/><k n="State" t="s"/><k n="Population" t="i"/></s><s t="_entity"><k n="_DisplayString" t="s"/><k n="Ticker symbol" t="s"/><k n="Price" t="d"/></s><s t="_entity"><k n="_DisplayString" t="s"/></s></rvStructures>`,
		},
	});
	await openLanding(page);
	await (
		await landingFileInput(page)
	).setInputFiles({
		name: 'cities.xlsx',
		mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
		buffer: Buffer.from(bytes),
	});
	const badge = (kind: string) => grid(page).locator(`.xg-data-badge[data-kind="${kind}"]`);
	const paths = new Set<string | null>();
	for (const kind of ['geography', 'stocks', 'generic']) {
		await expect(badge(kind)).toBeVisible();
		paths.add(await badge(kind).locator('path').getAttribute('d'));
	}
	expect(paths.size).toBe(3);
	const icon = grid(page).locator('.xg-ic');
	await expect(icon).toBeVisible();
	const conditionalBox = await icon.boundingBox();
	const stocksBox = await badge('stocks').boundingBox();
	expect(stocksBox!.x).toBeGreaterThanOrEqual(conditionalBox!.x + conditionalBox!.width);
	await grid(page).focus();
	await grid(page).press('Shift+F10');
	await editor(page).getByRole('menuitem', { name: 'Show Data Type Card' }).press('Enter');
	await expect(editor(page).getByRole('dialog', { name: 'Linked Data Type' })).toBeVisible();
	const card = editor(page).locator('[data-dialog="linked-data-card"]');
	await expect(card.getByText('Washington', { exact: true })).toBeVisible();
	await expect(card.getByText('Refresh is unavailable.', { exact: false })).toBeVisible();
	await page.keyboard.press('Escape');
	await goToCell(page, 'B1');
	await typeInActiveCell(page, '=A1.State');
	await expect(grid(page).getByRole('gridcell', { name: 'Washington', exact: true })).toBeVisible();
});
