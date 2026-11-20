// @vitest-environment jsdom
import { signatureStrings } from 'ooxml-core/i18n';
import { afterEach, describe, expect, it } from 'vitest';
import { DocxEditorElement, registerDocxEditor } from './index';
import { EDITOR_LOCALES } from './localization';
import { officeSigned, tampered, unsupportedMethod } from '../notices/__fixtures__/signed-packages';
import type { OfficeUiSignatureStatus } from '../notices/signature-status';

registerDocxEditor();
afterEach(() => document.body.replaceChildren());

const SIGNED = officeSigned('hello-world-office-2010-technical-preview.docx');

async function open(bytes: Uint8Array, locale = 'en') {
	const editor = document.createElement('docx-editor') as DocxEditorElement;
	editor.setAttribute('locale', locale);
	document.body.append(editor);
	await editor.load(bytes);
	const find = (id: string) =>
		editor.shadowRoot!.querySelector<OfficeUiSignatureStatus>(`[data-testid="${id}"]`);
	const label = (id: string) => find(id)?.shadowRoot?.querySelector('.label')?.textContent ?? '';
	return { editor, find, label };
}

describe('Word signature state', () => {
	it.each([
		['invalid', () => tampered(SIGNED, 'word/document.xml')],
		['unsupported', () => unsupportedMethod(SIGNED)],
	] as const)('never labels an %s signature "Signed"', async (state, make) => {
		const { find, label } = await open(await make());
		expect(find('signature-badge')?.hidden).toBe(false);
		expect(find('signature-badge')?.dataset.state).toBe(state);
		expect(label('signature-badge')).toBe(signatureStrings.en.states[state].label);
		expect(label('signature-badge')).not.toMatch(/^Signed\b/u);
		expect(find('signature-notice')?.hidden).toBe(false);
	});

	it('shows a verified-untrusted signature, its signer and the Info panel', async () => {
		const { find, label, editor } = await open(SIGNED);
		expect(label('signature-badge')).toBe('Signed, signer not verified');
		find('signature-badge')!.shadowRoot!.querySelector<HTMLButtonElement>('button')!.click();
		const panel = editor.shadowRoot!.querySelector<OfficeUiSignatureStatus>(
			'[data-testid="signature-panel"]',
		);
		expect(panel?.summary?.state).toBe('verified-untrusted');
		expect(panel?.shadowRoot?.querySelector('.signer')?.textContent).toBe('Test');
	});

	it('dismisses the load notice and shows nothing for an unsigned document', async () => {
		const { find } = await open(SIGNED);
		find('signature-notice')!.shadowRoot!.querySelector<HTMLButtonElement>('.dismiss')!.click();
		expect(find('signature-notice')?.hidden).toBe(true);
		expect(find('signature-badge')?.hidden).toBe(false);
		const blank = document.createElement('docx-editor') as DocxEditorElement;
		document.body.append(blank);
		const badge = blank.shadowRoot!.querySelector<HTMLElement>('[data-testid="signature-badge"]');
		expect(badge?.hidden).toBe(true);
		expect(
			blank.shadowRoot!.querySelector<HTMLElement>('[data-testid="signature-notice"]')?.hidden,
		).toBe(true);
	});

	it.each(EDITOR_LOCALES.filter((locale) => locale !== 'en'))(
		'words the badge in %s',
		async (locale) => {
			const { label } = await open(SIGNED, locale);
			expect(label('signature-badge')).toBe(
				signatureStrings[locale].states['verified-untrusted'].label,
			);
		},
	);
});
