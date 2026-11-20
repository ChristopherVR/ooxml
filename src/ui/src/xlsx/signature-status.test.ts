// @vitest-environment jsdom
import { SUPPORTED_OFFICE_LOCALES, signatureStrings } from 'ooxml-core/i18n';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import { officeSigned, tampered, unsupportedMethod } from '../notices/__fixtures__/signed-packages';
import type { OfficeUiSignatureStatus } from '../notices/signature-status';

const { XlsxEditorElement, defineXlsxEditor } = await import('./index');
beforeAll(() => defineXlsxEditor());
afterEach(() => document.body.replaceChildren());

const SIGNED = officeSigned('hello-world-signed.xlsx');

async function open(bytes: Uint8Array | undefined, locale = 'en') {
	const editor = document.createElement('xlsx-editor') as InstanceType<typeof XlsxEditorElement>;
	editor.setAttribute('locale', locale);
	document.body.append(editor);
	if (bytes) await editor.load(bytes, 'Signed.xlsx');
	else editor.newWorkbook();
	const find = (id: string) =>
		editor.shadowRoot!.querySelector<OfficeUiSignatureStatus>(`[data-testid="${id}"]`);
	const label = (id: string) => find(id)?.shadowRoot?.querySelector('.label')?.textContent ?? '';
	return { editor, find, label };
}

describe('Excel signature state', () => {
	it.each([
		['invalid', () => tampered(SIGNED, 'xl/workbook.xml')],
		['unsupported', () => unsupportedMethod(SIGNED)],
	] as const)('never labels an %s signature "Signed"', async (state, make) => {
		const { find, label } = await open(await make());
		expect(find('signature-badge')?.dataset.state).toBe(state);
		expect(label('signature-badge')).toBe(signatureStrings.en.states[state].label);
		expect(label('signature-badge')).not.toMatch(/^Signed\b/u);
		expect(find('signature-notice')?.hidden).toBe(false);
	});

	it('shows a verified-untrusted signature and its state on Info', async () => {
		const { find, label, editor } = await open(SIGNED);
		expect(label('signature-badge')).toBe('Signed, signer not verified');
		find('signature-notice')!.shadowRoot!.querySelector<HTMLButtonElement>('.view')!.click();
		const panel = editor.shadowRoot!.querySelector<OfficeUiSignatureStatus>(
			'[data-testid="signature-panel"]',
		);
		expect(panel?.summary?.state).toBe('verified-untrusted');
		expect(panel?.summary?.signatures[0]?.signer).toMatch(/CN=/u);
		expect(panel?.shadowRoot?.querySelector('.label')?.textContent).toBe(
			'Signed, signer not verified',
		);
	});

	it('shows nothing for an unsigned workbook', async () => {
		const { find } = await open(undefined);
		expect(find('signature-badge')?.hidden).toBe(true);
		expect(find('signature-notice')?.hidden).toBe(true);
	});

	it.each(SUPPORTED_OFFICE_LOCALES.filter((locale) => locale !== 'en'))(
		'words the badge in %s',
		async (locale) => {
			const { label } = await open(SIGNED, locale);
			expect(label('signature-badge')).toBe(
				signatureStrings[locale].states['verified-untrusted'].label,
			);
		},
	);

	it('keeps the signature state visible but inert while File is hidden', async () => {
		const { find, label, editor } = await open(SIGNED);
		const badge = find('signature-badge')!;
		const notice = find('signature-notice')!;
		const backstage = () =>
			(editor.shadowRoot!.querySelector('office-ui-backstage') as HTMLElement & { open: boolean })
				.open;
		editor.customization = { ribbon: { hiddenTabs: ['file'] } };
		await new Promise((resolve) => setTimeout(resolve));
		expect(badge.hidden).toBe(false);
		expect(label('signature-badge')).toBe('Signed, signer not verified');
		expect(badge.shadowRoot!.querySelector('button')).toBeNull();
		expect(badge.shadowRoot!.querySelector('.hidden-text')?.textContent).toContain(
			signatureStrings.en.heading,
		);
		expect(notice.hidden).toBe(false);
		expect(notice.shadowRoot!.querySelector('.view')).toBeNull();
		expect(notice.shadowRoot!.querySelector('.dismiss')).not.toBeNull();
		badge.dispatchEvent(new CustomEvent('office-signature-activate', { bubbles: true }));
		expect(backstage()).toBe(false);
		editor.customization = null;
		await new Promise((resolve) => setTimeout(resolve));
		badge.shadowRoot!.querySelector<HTMLButtonElement>('button')!.click();
		expect(backstage()).toBe(true);
	});
});
