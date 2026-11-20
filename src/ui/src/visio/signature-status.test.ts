// @vitest-environment jsdom
import { signatureStrings } from 'ooxml-core/i18n';
import { demoDocument } from 'ooxml-core/visio/ui';
import { afterEach, describe, expect, it } from 'vitest';
import { testSigned, tampered, unsupportedMethod } from '../notices/__fixtures__/signed-packages';
import type { OfficeUiSignatureStatus } from '../notices/signature-status';
import { mountViewer } from './binding';
import { VISIO_LOCALES } from './ribbon-locale';

afterEach(() => document.body.replaceChildren());

const SIGNED = testSigned('signed-drawing.vsdx');

async function open(bytes: Uint8Array | undefined, locale = 'en') {
	const host = document.createElement('div');
	document.body.append(host);
	const viewer = mountViewer(host, { document: demoDocument });
	viewer.element.setAttribute('locale', locale);
	if (bytes) await viewer.element.load(bytes);
	const root = viewer.element.shadowRoot!;
	const find = (id: string) => root.querySelector<OfficeUiSignatureStatus>(`[data-testid="${id}"]`);
	const label = (id: string) => find(id)?.shadowRoot?.querySelector('.label')?.textContent ?? '';
	return { viewer, find, label };
}

describe('Visio signature state', () => {
	it.each([
		['invalid', () => tampered(SIGNED, 'visio/pages/page1.xml')],
		['unsupported', () => unsupportedMethod(SIGNED)],
	] as const)(
		'never labels an %s signature "Signed"',
		async (state, make) => {
			const { find, label, viewer } = await open(await make());
			expect(find('signature-badge')?.dataset.state).toBe(state);
			expect(label('signature-badge')).toBe(signatureStrings.en.states[state].label);
			expect(label('signature-badge')).not.toMatch(/^Signed\b/u);
			expect(find('signature-notice')?.hidden).toBe(false);
			viewer.destroy();
		},
		60_000,
	);

	it('shows a verified-untrusted drawing, its signer on Info, and nothing unsigned', async () => {
		const blank = await open(undefined);
		expect(blank.find('signature-badge')?.hidden).toBe(true);
		blank.viewer.destroy();
		const { find, label, viewer } = await open(SIGNED);
		expect(label('signature-badge')).toBe('Signed, signer not verified');
		find('signature-badge')!.shadowRoot!.querySelector<HTMLButtonElement>('button')!.click();
		const panel = find('signature-panel')!;
		expect(panel.summary?.state).toBe('verified-untrusted');
		expect(panel.shadowRoot?.querySelector('.signer')?.textContent).toBe('OOXML Test Signer');
		viewer.destroy();
	}, 60_000);

	it('words the badge in every Visio locale', async () => {
		const { viewer, label } = await open(SIGNED);
		for (const locale of VISIO_LOCALES) {
			viewer.element.setAttribute('locale', locale);
			expect(label('signature-badge')).toBe(
				signatureStrings[locale].states['verified-untrusted'].label,
			);
		}
		viewer.destroy();
	}, 60_000);
});
