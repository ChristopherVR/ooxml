import { afterEach, describe, expect, it } from 'vitest';
import { defineRibbon } from '../controls';
import { applyHostRibbonTabs } from './host-ribbon-tabs';

afterEach(() => document.body.replaceChildren());

type Ribbon = HTMLElement & {
	noFile: boolean;
	selected: string;
	fileButton(): HTMLButtonElement | null;
	focusFile(): void;
	setTabHidden(tab: string, reason: string, hidden: boolean): void;
	updateComplete: Promise<unknown>;
};

function ribbon(): Ribbon {
	defineRibbon();
	const element = document.createElement('office-ui-ribbon') as Ribbon;
	element.setAttribute('file-keytip', 'F');
	for (const id of ['home', 'insert', 'table']) {
		const panel = document.createElement('div');
		panel.dataset.ribbonTab = id;
		panel.dataset.label = id;
		element.append(panel);
	}
	const addIn = document.createElement('div');
	addIn.dataset.ribbonTab = 'home-addin';
	addIn.dataset.addIn = '';
	element.append(addIn);
	document.body.append(element);
	return element;
}
const tabs = async (element: Ribbon) => {
	await new Promise((resolve) => setTimeout(resolve));
	await element.updateComplete;
	return [...element.shadowRoot!.querySelectorAll<HTMLElement>('[role="tab"]')].map(
		(tab) => tab.dataset.tab,
	);
};

describe('applyHostRibbonTabs', () => {
	it('hides File, its key tip and named native tabs, and shows them again', async () => {
		const element = ribbon();
		element.setTabHidden('table', 'contextual', true);
		applyHostRibbonTabs(element, new Set(['file', 'home', 'home-addin']));
		expect(element.noFile).toBe(true);
		expect(element.fileButton()).toBeNull();
		const file = element.shadowRoot!.querySelector<HTMLElement>('.file')!;
		expect(file.hidden).toBe(true);
		expect(file.hasAttribute('data-keytip')).toBe(false);
		// Host add-in tabs are not native tabs; the contextual reason is kept apart.
		expect(await tabs(element)).toEqual(['insert', 'home-addin']);
		expect(element.selected).toBe('insert');
		element.focusFile();
		expect(element.shadowRoot!.activeElement?.getAttribute('data-tab')).toBe('insert');
		applyHostRibbonTabs(element, new Set());
		expect(element.noFile).toBe(false);
		expect(element.fileButton()).toBe(file);
		expect(file.getAttribute('data-keytip')).toBe('F');
		expect(await tabs(element)).toEqual(['home', 'insert', 'home-addin']);
	});
});
