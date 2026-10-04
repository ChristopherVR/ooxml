import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import type { OfficeRibbonGroupView } from './controls.js';
import { registerOfficeUi } from './index.js';

// Adapted from pptx-viewer's ribbon-section test (`packages/shared/src/web-components/ribbon-section.test.ts`).
beforeAll(() => registerOfficeUi());
afterEach(() => document.body.replaceChildren());

type Section = HTMLElement & { groups: readonly OfficeRibbonGroupView[] };

const groups = (spelling: boolean): OfficeRibbonGroupView[] => [
	{
		id: 'proofing',
		label: 'Proofing',
		commands: [
			{ id: 'spelling', label: 'Spelling', icon: 'check', pressed: spelling, active: spelling },
			{ id: 'cut', label: 'Cut', icon: 'cut', column: 1, size: 'small' },
			{ id: 'copy', label: 'Copy', icon: 'copy', column: 1, size: 'small', disabled: true },
		],
	},
	{ id: 'comments', label: 'Comments', commands: [{ id: 'new', label: 'New', icon: 'message' }] },
];

function make(): Section {
	const section = document.createElement('office-ui-ribbon-section') as Section;
	document.body.append(section);
	return section;
}

describe('office-ui-ribbon-section', () => {
	it('renders groups and commands with attributes and stacked columns', () => {
		const section = make();
		section.groups = groups(false);
		const [proofing, comments] = [...section.children] as HTMLElement[];
		expect(proofing!.tagName).toBe('OFFICE-UI-RIBBON-GROUP');
		expect(proofing!.getAttribute('data-ribbon-group')).toBe('proofing');
		expect(comments!.getAttribute('label')).toBe('Comments');
		const spelling = section.querySelector('[command="spelling"]')!;
		expect(spelling.getAttribute('pressed')).toBe('false');
		expect(spelling.getAttribute('title')).toBe('Spelling');
		const column = section.querySelector('[command="cut"]')!.parentElement!;
		expect(column.tagName).toBe('DIV');
		expect([...column.children].map((c) => c.getAttribute('command'))).toEqual(['cut', 'copy']);
		expect(section.querySelector('[command="copy"]')!.hasAttribute('disabled')).toBe(true);
	});

	it('keeps focus and nodes across updates, remounts and independent instances', () => {
		const first = make();
		const second = make();
		first.groups = groups(false);
		second.groups = groups(false);
		const command = first.querySelector<HTMLElement>('[command="spelling"]')!;
		const button = command.shadowRoot!.querySelector('button')!;
		button.focus();
		first.groups = groups(true);
		expect(document.activeElement).toBe(command);
		expect(button.getAttribute('aria-pressed')).toBe('true');
		expect(second.querySelector('[command="spelling"]')!.hasAttribute('active')).toBe(false);
		first.remove();
		document.body.append(first);
		expect(first.querySelector('[command="spelling"]')).toBe(command);
		first.groups = [groups(false)[1]!];
		expect(first.querySelector('[command="spelling"]')).toBeNull();
		first.groups = [];
		expect(first.childElementCount).toBe(0);
	});

	it('lets commands emit their own events and products swap tags and ids', () => {
		const section = make();
		section.groups = groups(false);
		const seen = vi.fn();
		section.addEventListener('office-command', (e) => seen((e as CustomEvent).detail));
		section
			.querySelector<HTMLElement>('[command="new"]')!
			.shadowRoot!.querySelector('button')!
			.click();
		expect(seen).toHaveBeenCalledWith({ command: 'new' });

		const Base = customElements.get('office-ui-ribbon-section') as unknown as { new (): Section };
		class Product extends Base {
			static commandIdAttribute = 'data-ribbon-control';
		}
		customElements.define('product-ribbon-section', Product as unknown as CustomElementConstructor);
		const product = document.createElement('product-ribbon-section') as Section;
		document.body.append(product);
		product.groups = groups(false);
		expect(product.querySelector('[data-ribbon-control="cut"]')).not.toBeNull();
	});
});
