import { registerOfficeUi } from '../index';

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

	it('takes a translated name for its close button', () => {
		const { dialog } = mount('close-label="Fermer"');
		expect(dialog.shadowRoot!.querySelector('.close')!.getAttribute('aria-label')).toBe('Fermer');
	});

	it('moves focus in on open and restores it on close', () => {
		const { dialog, opener } = mount();
		dialog.show();
		expect(document.activeElement?.id).toBe('field');
		dialog.open = false;
		expect(document.activeElement).toBe(opener);
	});

	it('Escape emits a cancelable office-dialog-close and closes unless prevented', () => {
		const { dialog, opener } = mount();
		dialog.show();
		const reasons: string[] = [];
		const closed = vi.fn(() => {
			expect(dialog.hasAttribute('open')).toBe(false);
			expect(document.activeElement).toBe(opener);
		});
		dialog.addEventListener('office-dialog-closed', closed);
		let prevent = true;
		dialog.addEventListener('office-dialog-close', (e) => {
			reasons.push((e as CustomEvent).detail.reason);
			if (prevent) e.preventDefault();
		});
		key(dialog.querySelector('#field')!, 'Escape');
		expect(dialog.open).toBe(true);
		expect(closed).not.toHaveBeenCalled();
		prevent = false;
		key(dialog.querySelector('#field')!, 'Escape');
		expect(dialog.open).toBe(false);
		expect(reasons).toEqual(['escape', 'escape']);
		expect(closed).toHaveBeenCalledOnce();
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

	// Tab order only exists in a real browser: e2e/a11y/dialog.spec.ts and e2e/docx/dialog-focus.spec.ts.
	it('closes on Escape when focus has ended up outside', () => {
		const { dialog, opener } = mount('open');
		opener.focus();
		key(opener, 'Escape');
		expect(dialog.open).toBe(false);
	});

	it('gives the keyboard to the newest of two open dialogs', () => {
		const { dialog } = mount('open');
		const inner = document.createElement('office-ui-dialog') as Dialog;
		inner.setAttribute('heading', 'Inner');
		inner.append(document.createElement('input'));
		document.body.append(inner);
		inner.show();
		key(document.body, 'Escape');
		expect(inner.open).toBe(false);
		expect(dialog.open).toBe(true);
		key(document.body, 'Escape');
		expect(dialog.open).toBe(false);
	});

	it('hides the footer without footer content', () => {
		document.body.innerHTML = '<office-ui-dialog open heading="x">body</office-ui-dialog>';
		const dialog = document.querySelector('office-ui-dialog') as HTMLElement;
		expect(dialog.shadowRoot!.querySelector('footer')!.hidden).toBe(true);
	});
});
