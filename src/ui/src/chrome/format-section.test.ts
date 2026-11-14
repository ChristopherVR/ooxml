import { afterEach, describe, expect, it, vi } from 'vitest';
import { defineFormatSection, type OfficeUiFormatSection } from './format-section';

afterEach(() => document.body.replaceChildren());

function make(): OfficeUiFormatSection {
	defineFormatSection();
	const section = document.createElement('office-ui-format-section') as OfficeUiFormatSection;
	section.label = 'Fill';
	section.append(document.createTextNode('Body'));
	document.body.append(section);
	return section;
}

describe('office-ui-format-section', () => {
	it('renders a native disclosure that follows the open property', () => {
		const section = make();
		const details = section.shadowRoot!.querySelector('details')!;
		expect(section.shadowRoot!.querySelector('summary')!.textContent).toBe('Fill');
		expect(details.open).toBe(false);
		section.open = true;
		expect(details.open).toBe(true);
		expect(section.hasAttribute('open')).toBe(true);
		expect(section.textContent).toBe('Body');
	});

	it('toggles from the heading and emits office-section-toggle', () => {
		const section = make();
		const onToggle = vi.fn((event: Event) => (event as CustomEvent<{ open: boolean }>).detail);
		section.addEventListener('office-section-toggle', onToggle);
		const summary = section.shadowRoot!.querySelector('summary')!;
		summary.click();
		expect(onToggle).toHaveReturnedWith({ open: true });
		expect(section.open).toBe(true);
		summary.click();
		expect(onToggle).toHaveLastReturnedWith({ open: false });
		expect(section.open).toBe(false);
	});

	it('stays as it was when the toggle is cancelled', () => {
		const section = make();
		section.addEventListener('office-section-toggle', (event) => event.preventDefault());
		section.shadowRoot!.querySelector('summary')!.click();
		expect(section.open).toBe(false);
	});

	it('shows a mixed badge and focuses its heading', () => {
		const section = make();
		expect(section.shadowRoot!.querySelector('.badge')).toBeNull();
		section.mixed = true;
		expect(section.shadowRoot!.querySelector('.badge')!.textContent).toBe('Mixed');
		section.mixedLabel = 'Gemischt';
		expect(section.shadowRoot!.querySelector('.badge')!.textContent).toBe('Gemischt');
		section.focus();
		expect(section.shadowRoot!.activeElement).toBe(section.shadowRoot!.querySelector('summary'));
	});
});
