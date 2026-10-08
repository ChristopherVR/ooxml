import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import { registerOfficeUi } from '../index';
import type { OfficeUiRibbonActions } from './ribbon-actions';

beforeAll(() => registerOfficeUi());
afterEach(() => document.body.replaceChildren());

function make(): OfficeUiRibbonActions {
	const element = document.createElement('office-ui-ribbon-actions') as OfficeUiRibbonActions;
	document.body.append(element);
	return element;
}

const parts = (element: OfficeUiRibbonActions) =>
	[...element.shadowRoot!.children]
		.filter((child) => child.tagName !== 'STYLE')
		.map((child) => child.getAttribute('part'));

describe('office-ui-ribbon-actions', () => {
	it("draws Office's order: editing mode, Comments, Share", () => {
		const element = make();
		element.modes = [
			{ value: 'editing', label: 'Editing' },
			{ value: 'viewing', label: 'Viewing' },
		];
		element.mode = 'viewing';
		expect(parts(element)).toEqual(['mode', 'comments', 'share']);
		const select = element.control('mode-select') as HTMLSelectElement;
		expect(select.getAttribute('aria-label')).toBe('Editing mode');
		expect([...select.options].map((option) => option.textContent)).toEqual(['Editing', 'Viewing']);
		expect(select.value).toBe('viewing');
		expect(element.control('comments')!.textContent).toBe('Comments');
		expect(element.control('share')!.textContent).toBe('Share');
		expect(element.control('comments')!.getAttribute('aria-label')).toBe('Comments');
	});

	it('hides the selector without modes and either button on request', () => {
		const element = make();
		expect(element.control('mode-select')!.closest('[part="mode"]')!.hasAttribute('hidden')).toBe(
			true,
		);
		element.noComments = true;
		element.setAttribute('no-share', '');
		expect((element.control('comments') as HTMLButtonElement).hidden).toBe(true);
		expect((element.control('share') as HTMLButtonElement).hidden).toBe(true);
	});

	it('is controlled: emits the picked mode and keeps the host mode until redrawn', () => {
		const element = make();
		element.modes = [
			{ value: 'editing', label: 'Editing' },
			{ value: 'viewing', label: 'Viewing' },
		];
		element.mode = 'editing';
		const picked: string[] = [];
		element.addEventListener('office-ribbon-mode', (event) =>
			picked.push((event as CustomEvent<{ mode: string }>).detail.mode),
		);
		const select = element.control('mode-select') as HTMLSelectElement;
		select.value = 'viewing';
		select.dispatchEvent(new Event('change', { bubbles: true }));
		expect(picked).toEqual(['viewing']);
		expect(select.value).toBe('editing');
		element.mode = 'viewing';
		expect(select.value).toBe('viewing');
	});

	it('reports Comments and Share clicks and reflects pressed states, labels and the count', () => {
		const element = make();
		const events: string[] = [];
		for (const type of ['office-ribbon-comments', 'office-ribbon-share'])
			element.addEventListener(type, () => events.push(type));
		const comments = element.control('comments') as HTMLButtonElement;
		const share = element.control('share') as HTMLButtonElement;
		comments.click();
		share.click();
		expect(events).toEqual(['office-ribbon-comments', 'office-ribbon-share']);
		expect(comments.getAttribute('aria-pressed')).toBe('false');
		element.commentsPressed = true;
		element.sharePressed = true;
		element.commentsTitle = 'Show comments';
		element.commentsCount = 3;
		element.shareLabel = 'Partager';
		expect(comments.getAttribute('aria-pressed')).toBe('true');
		expect(share.getAttribute('aria-pressed')).toBe('true');
		expect(comments.title).toBe('Show comments (3)');
		// The badge is decoration: the accessible name stays the label.
		expect(comments.getAttribute('aria-label')).toBe('Comments');
		expect(comments.querySelector('.badge')!.textContent).toBe('3');
		expect(share.textContent).toBe('Partager');
	});
});
