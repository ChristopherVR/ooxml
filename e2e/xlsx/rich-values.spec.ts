import { expect, test, type Page } from '@playwright/test';
import type { XlsxEditorElement } from '../../src/ui/src/xlsx/component';
import { RICH_METADATA_EXT, RICH_NS, RICH_PARTS } from '../../src/core/xlsx/rich-values';
import { miniPackage, R, ws, X } from '../../src/core/xlsx/write/mini-package-fixtures';
import {
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

test('uses the host IMAGE provider on the already open editor', async ({ page }) => {
	await openLanding(page);
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
	await expect(grid(page).getByRole('img', { name: 'Host image' })).toBeVisible();
	await editor(page).evaluate((element) => {
		const xlsx = element as XlsxEditorElement;
		xlsx.imageProvider = undefined;
		xlsx.calculateNow();
	});
	await expect(grid(page).getByRole('img', { name: 'Host image' })).toHaveCount(0);
	await expect(grid(page).getByText('#BLOCKED!', { exact: true })).toBeVisible();
});

test('disables in-cell picture editing with a reason while sharing', async ({ page }) => {
	await openLanding(page);
	await newWorkbook(page);
	await editor(page).evaluate(
		(element, roomId) =>
			(element as XlsxEditorElement).startCollaboration({ roomId, user: { name: 'Ada' } }),
		`pictures-${Date.now()}`,
	);
	await openPictures(page);
	const command = editor(page).getByRole('menuitem', { name: 'Place in Cell', exact: true });
	await expect(command).toBeDisabled();
	await expect(command).toHaveAccessibleDescription(
		'In-cell picture editing is unavailable while sharing this workbook.',
	);
	await page.keyboard.press('Escape');
	await editor(page).evaluate((element) => (element as XlsxEditorElement).stopCollaboration());
	await openPictures(page);
	await expect(command).toBeEnabled();
});

test('shows a cached data type icon and opens its card from the keyboard', async ({ page }) => {
	const { data, structures } = RICH_PARTS;
	const bytes = await miniPackage({
		sheets: [
			{ name: 'Cities', xml: ws('<row r="1"><c r="A1" vm="1" t="e"><v>#VALUE!</v></c></row>') },
		],
		workbookRels: `<Relationship Id="meta" Type="${R}/sheetMetadata" Target="metadata.xml"/><Relationship Id="data" Type="${data.rel}" Target="richData/${data.file}"/><Relationship Id="structures" Type="${structures.rel}" Target="richData/${structures.file}"/>`,
		parts: {
			'xl/metadata.xml': `<metadata xmlns="${X}" xmlns:r="${RICH_NS}"><metadataTypes count="1"><metadataType name="XLRICHVALUE" minSupportedVersion="120000"/></metadataTypes><futureMetadata name="XLRICHVALUE" count="1"><bk><extLst><ext uri="${RICH_METADATA_EXT}"><r:rvb i="0"/></ext></extLst></bk></futureMetadata><valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata></metadata>`,
			[`xl/richData/${data.file}`]: `<rvData xmlns="${RICH_NS}" count="1"><rv s="0"><v>Seattle</v><v>Washington</v></rv></rvData>`,
			[`xl/richData/${structures.file}`]: `<rvStructures xmlns="${RICH_NS}" count="1"><s t="_entity"><k n="_DisplayString" t="s"/><k n="State" t="s"/></s></rvStructures>`,
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
	await expect(
		grid(page).locator('.xg-c').filter({ hasText: 'Seattle' }).locator('.xg-ic'),
	).toBeVisible();
	await grid(page).focus();
	await grid(page).press('Shift+F10');
	await editor(page).getByRole('menuitem', { name: 'Show Data Type Card' }).press('Enter');
	await expect(editor(page).getByRole('dialog', { name: 'Linked Data Type' })).toBeVisible();
	const card = editor(page).locator('[data-dialog="linked-data-card"]');
	await expect(card.getByText('Washington', { exact: true })).toBeVisible();
	await expect(card.getByText('Refresh is unavailable.', { exact: false })).toBeVisible();
});
