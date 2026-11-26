// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createRibbon } from './ribbon';
import './component';
import { findLocalizedControl } from './localization';
import { must, ribbonTab } from './test-support';

afterEach(() => document.body.replaceChildren());

describe('Word native tab surfaces', () => {
	it('places the existing page background commands on Design', () => {
		const ribbon = createRibbon();
		expect(ribbon.querySelector('[data-panel="Design"] [aria-label="Watermark"]')).not.toBeNull();
		expect(
			ribbon.querySelector('[data-panel="Design"] [aria-label="Page borders"]'),
		).not.toBeNull();
		expect(ribbon.querySelector('[data-panel="Layout"] [data-label="Page background"]')).toBeNull();
	});

	it('enables Mailings commands only when the editor can use them and localizes the reasons', async () => {
		const editor = document.createElement('docx-editor');
		document.body.append(editor);
		const root = must(editor.shadowRoot);
		const panel = must(root.querySelector('[data-panel="Mailings"]'));
		(await ribbonTab(root, 'mailings')).click();
		await vi.waitFor(() => expect(panel.childElementCount).toBeGreaterThan(0), { timeout: 10_000 });
		const button = (label: string) => must(findLocalizedControl<HTMLButtonElement>(panel, label));
		expect(button('Start Mail Merge').disabled).toBe(false);
		expect(button('Select Recipients').disabled).toBe(false);
		expect(button('Preview Results').disabled).toBe(true);
		expect(button('Preview Results').title).toBe('Include at least one recipient');
		expect(button('Insert Merge Field').title).toBe('Select a recipient source first');
		expect(button('Envelopes').disabled).toBe(false);
		expect(button('Labels').disabled).toBe(false);
		expect(button('Update Labels').disabled).toBe(true);
		editor.locale = 'fr';
		expect(button('Update Labels').title).toBe(
			'Commencez par une fusion d’étiquettes avec une grille',
		);
		expect(button('Preview Results').title).toBe('Incluez au moins un destinataire');
		editor.readOnly = true;
		expect(button('Start Mail Merge').disabled).toBe(true);
	});
});
