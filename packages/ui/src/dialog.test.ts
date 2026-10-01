import { registerOfficeUi } from './index.js';

beforeAll(() => registerOfficeUi());
afterEach(() => document.body.replaceChildren());

type Dialog = HTMLElement & { open: boolean; show(): void; close(): void };
const key = (el: Element, k: string, init: KeyboardEventInit = {}) =>
	el.dispatchEvent(
		new KeyboardEvent('keydown', {
			key: k,
			bubbles: true,
			composed: true,
			cancelable: true,
			...init,
		}),
	);

function mount(attrs = ''): {
	dialog: Dialog;
	opener: HTMLButtonElement;
	ok: HTMLButtonElement;
	cancel: HTMLButtonElement;
} {
	document.body.innerHTML =
		'<button id="opener">open</button>' +
		`<office-ui-dialog heading="Options" ${attrs}><input id="field" />` +
		'<button slot="footer" id="cancel">Cancel</button><button slot="footer" id="ok">OK</button></office-ui-dialog>';
	const opener = document.getElementById('opener') as HTMLButtonElement;
	opener.focus();
	return {
		dialog: document.querySelector('office-ui-dialog') as Dialog,
		opener,
		ok: document.getElementById('ok') as HTMLButtonElement,
		cancel: document.getElementById('cancel') as HTMLButtonElement,
	};
}

describe('office-ui-dialog', () => {
	it('exposes dialog semantics labelled by its heading', () => {
		const { dialog } = mount();
		const box = dialog.shadowRoot!.querySelector('[role="dialog"]')!;
		expect(box.getAttribute('aria-modal')).toBe('true');
		const labelled = dialog.shadowRoot!.getElementById(box.getAttribute('aria-labelledby')!)!;
		expect(labelled.textContent).toBe('Options');
		dialog.setAttribute('heading', 'Page setup');
		expect(labelled.textContent).toBe('Page setup');
	});

	it('moves focus in on open and restores it on close', () => {
		const { dialog, opener } = mount();
		dialog.show();
		expect(document.activeElement?.id).toBe('field');
		dialog.open = false;
		expect(document.activeElement).toBe(opener);
	});

	it('Escape emits a cancelable office-dialog-close and closes unless prevented', () => {
		const { dialog } = mount('open');
		const reasons: string[] = [];
		let prevent = true;
		dialog.addEventListener('office-dialog-close', (e) => {
			reasons.push((e as CustomEvent).detail.reason);
			if (prevent) e.preventDefault();
		});
		key(dialog.querySelector('#field')!, 'Escape');
		expect(dialog.open).toBe(true);
		prevent = false;
		key(dialog.querySelector('#field')!, 'Escape');
		expect(dialog.open).toBe(false);
		expect(reasons).toEqual(['escape', 'escape']);
	});

	it('close button and backdrop request closing; dismissible=false ignores Escape and backdrop', () => {
		const { dialog } = mount('open dismissible="false"');
		const reasons: string[] = [];
		dialog.addEventListener('office-dialog-close', (e) =>
			reasons.push((e as CustomEvent).detail.reason),
		);
		key(dialog.querySelector('#field')!, 'Escape');
		(dialog.shadowRoot!.querySelector('.backdrop') as HTMLElement).click();
		expect(dialog.open).toBe(true);
		(dialog.shadowRoot!.querySelector('.close') as HTMLElement).click();
		expect(dialog.open).toBe(false);
		expect(reasons).toEqual(['close-button']);
	});

	it('traps Tab: from the last light-DOM control to the close button and around', () => {
		const { dialog, ok } = mount('open');
		const close = dialog.shadowRoot!.querySelector('.close') as HTMLElement;
		const field = dialog.querySelector('#field') as HTMLElement;
		ok.focus();
		const forward = key(ok, 'Tab');
		// Wrapping is only prevented at the ends; ok is last light-DOM item, the close button is the real last.
		expect(forward).toBe(true);
		close.focus();
		expect(key(close, 'Tab')).toBe(false);
		expect(document.activeElement).toBe(field);
		expect(key(field, 'Tab', { shiftKey: true })).toBe(false);
	});

	it('hides the footer without footer content', () => {
		document.body.innerHTML = '<office-ui-dialog open heading="x">body</office-ui-dialog>';
		const dialog = document.querySelector('office-ui-dialog') as HTMLElement;
		expect(dialog.shadowRoot!.querySelector('footer')!.hidden).toBe(true);
	});
});
