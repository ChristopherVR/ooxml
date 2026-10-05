import { afterEach, describe, expect, it } from 'vitest';
import { defineSelect } from './select.js';

afterEach(() => document.body.replaceChildren());

describe('office-ui-select relabelling', () => {
	it('repaints the trigger when aria-label changes after the first render', () => {
		defineSelect(customElements);
		const select = document.createElement('office-ui-select');
		select.setAttribute('aria-label', 'Font size');
		document.body.append(select);
		const trigger = () => select.shadowRoot!.querySelector('button[role="combobox"]')!;
		expect(trigger().getAttribute('aria-label')).toBe('Font size');
		select.setAttribute('aria-label', 'Schriftgröße');
		expect(trigger().getAttribute('aria-label')).toBe('Schriftgröße');
	});
});
