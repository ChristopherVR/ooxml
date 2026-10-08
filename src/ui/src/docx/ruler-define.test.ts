// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { createRuler, updateRuler } from './ruler';

// Deliberately no `registerOfficeUi()`: the editor never calls it, so the ruler must register
// `office-ui-ruler` itself or it renders as an unknown element without markers.
describe('createRuler', () => {
	it('registers and upgrades the shared ruler element', () => {
		const ruler = createRuler();
		const OfficeRuler = customElements.get('office-ui-ruler');
		expect(OfficeRuler).toBeDefined();
		expect(ruler).toBeInstanceOf(OfficeRuler!);
		document.body.append(ruler);
		updateRuler(
			ruler,
			{
				pageWidth: 816,
				marginLeft: 96,
				marginRight: 96,
				indentLeft: 0,
				indentRight: 0,
				firstLine: 0,
			},
			1,
		);
		expect(ruler.getAttribute('role')).toBe('img');
		expect(ruler.getAttribute('aria-label')).toBe('Ruler');
		expect(ruler.shadowRoot?.querySelector('[data-marker="left"]')).not.toBeNull();
		ruler.remove();
	});
});
