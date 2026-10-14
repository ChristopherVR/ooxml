// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { createRibbon, setRibbonLocale } from './ribbon';
import { refreshEditorControls } from './editor-controls';
import { createDocument } from 'ooxml-core/docx';

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

	it('keeps unavailable mail merge commands disabled after refresh and locale changes', () => {
		const ribbon = createRibbon();
		refreshEditorControls(ribbon, undefined, createDocument(), false, false, 'en', 'en-US');
		const mailings = ribbon.querySelector('[data-panel="Mailings"]')!;
		const commands = [...mailings.querySelectorAll<HTMLButtonElement>('button')];
		expect(commands.length).toBeGreaterThan(10);
		expect(commands.every((command) => command.disabled && !command.dataset.action)).toBe(true);
		expect(commands[0]!.title).toBe('Not available in this editor');
		setRibbonLocale(ribbon, 'fr');
		expect(commands[0]!.title).toBe('Indisponible dans cet éditeur');
		expect(commands.every((command) => command.disabled)).toBe(true);
	});
});
