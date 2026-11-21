import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { expect, test, type Page } from '@playwright/test';
import { fileInput, openSample } from './helpers';

// A Word-signed sample (Apache POI test data): its signature verifies, its signer is not trusted.
const SIGNED = readFileSync(
	new URL(
		'../../src/core/opc/signature/__fixtures__/office/hello-world-office-2010-technical-preview.docx',
		import.meta.url,
	),
);

async function rewrite(change: (zip: JSZip) => Promise<void>): Promise<Buffer> {
	const zip = await JSZip.loadAsync(SIGNED);
	await change(zip);
	return zip.generateAsync({ type: 'nodebuffer' });
}
/** A signed part changed after signing. */
const tampered = () =>
	rewrite(async (zip) => {
		const xml = await zip.file('word/document.xml')!.async('string');
		zip.file('word/document.xml', xml.replace('<w:body>', '<w:body><!-- changed -->'));
	});
/** A signature method the verifier does not implement. */
const unsupported = () =>
	rewrite(async (zip) => {
		for (const name of Object.keys(zip.files).filter((path) =>
			/^_xmlsignatures\/.+\.xml$/u.test(path),
		)) {
			const xml = await zip.file(name)!.async('string');
			zip.file(
				name,
				xml.replace(
					/(SignatureMethod Algorithm=")[^"]+"/u,
					'$1http://www.w3.org/2001/04/xmldsig-more#ecdsa-sha256"',
				),
			);
		}
	});

async function open(page: Page, buffer: Buffer) {
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'signed.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer,
	});
	await expect(page.locator('docx-editor .dve-titlebar .file .name')).toContainText('signed');
}

const badge = (page: Page) => page.locator('docx-editor [data-testid="signature-badge"]');
const notice = (page: Page) => page.locator('docx-editor [data-testid="signature-notice"]');

test('Word shows each signature state from a real file, and nothing unsigned', async ({ page }) => {
	await openSample(page);
	await expect(badge(page)).toBeHidden();
	await expect(notice(page)).toBeHidden();

	const cases: [string, Buffer, string][] = [
		['verified-untrusted', SIGNED, 'Signed, signer not verified'],
		['invalid', await tampered(), 'Invalid signature'],
		['unsupported', await unsupported(), 'Signature not verified'],
	];
	for (const [state, buffer, label] of cases) {
		await open(page, buffer);
		await expect(badge(page)).toHaveAttribute('data-state', state);
		await expect(badge(page).getByRole('button')).toHaveText(label);
		await expect(notice(page)).toBeVisible();
		await expect(notice(page)).toContainText('Saving changes removes the signatures');
	}

	// Wait for the reopened file before using its notice. Every fixture has the same title.
	await open(page, SIGNED);
	await expect(badge(page)).toHaveAttribute('data-state', 'verified-untrusted');
	await notice(page).getByRole('button', { name: 'View signatures' }).click();
	const panel = page.locator('docx-editor [data-testid="signature-panel"]');
	await expect(panel).toContainText('Signed, signer not verified');
	await expect(panel).toContainText('Test');
});

test('a hidden File tab keeps the signature state on screen, without actions', async ({ page }) => {
	await openSample(page);
	await page.locator('docx-editor').evaluate((element) => {
		(element as HTMLElement & { customization: unknown }).customization = {
			ribbon: { hiddenTabs: ['file'] },
		};
	});
	await open(page, SIGNED);
	await expect(badge(page)).toBeVisible();
	await expect(badge(page)).toContainText('Signed, signer not verified');
	await expect(badge(page).getByRole('button')).toHaveCount(0);
	await expect(notice(page)).toBeVisible();
	await expect(notice(page).getByRole('button', { name: 'View signatures' })).toHaveCount(0);
	await expect(page.locator('docx-editor .dve-backstage')).toBeHidden();
});
