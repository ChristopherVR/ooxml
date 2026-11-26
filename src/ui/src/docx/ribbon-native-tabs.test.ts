// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { createRibbon, loadTablePanels, setRibbonLocale } from './ribbon';
import './component';
import { findLocalizedControl } from './localization';
import { must } from './test-support';

afterEach(() => document.body.replaceChildren());

describe('Word native tab surfaces', () => {
	it('loads localized table controls on demand and translates them after a locale change', async () => {
		const ribbon = createRibbon('fr');
		expect(ribbon.querySelector('[data-table-command]')).toBeNull();
		await loadTablePanels(ribbon);
		expect(
			ribbon.querySelector('[data-ribbon-tab="table-design"]')?.getAttribute('data-label'),
		).toBe('Création de tableau');
		expect(ribbon.querySelector('[data-look-flag="firstRow"]')?.getAttribute('aria-label')).toBe(
			'Ligne d’en-tête',
		);
		expect(ribbon.querySelector('[data-table-command="formula"]')?.getAttribute('aria-label')).toBe(
			'Formule',
		);
		setRibbonLocale(ribbon, 'de');
		expect(ribbon.querySelector('[data-table-command="formula"]')?.getAttribute('aria-label')).toBe(
			'Formel',
		);
	});

	it('places the existing page background commands on Design', () => {
		const ribbon = createRibbon();
		expect(ribbon.querySelector('[data-panel="Design"] [aria-label="Watermark"]')).not.toBeNull();
		expect(
			ribbon.querySelector('[data-panel="Design"] [aria-label="Page borders"]'),
		).not.toBeNull();
		expect(ribbon.querySelector('[data-panel="Layout"] [data-label="Page background"]')).toBeNull();
	});

	it('enables Mailings commands only when the editor can use them and localizes the reasons', () => {
		const editor = document.createElement('docx-editor');
		document.body.append(editor);
		const root = must(editor.shadowRoot);
		const panel = must(root.querySelector('[data-panel="Mailings"]'));
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
