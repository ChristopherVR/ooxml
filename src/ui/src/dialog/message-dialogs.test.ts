import { alertDialog, confirmDialog, promptDialog } from '../controls';

afterEach(() => {
	document.body.replaceChildren();
});

const dialogs = (root: ParentNode = document) =>
	Array.from(root.querySelectorAll('office-ui-message-dialog'));
const shadowOf = (el: Element) => el.shadowRoot as ShadowRoot;
const button = (el: Element, cls: string) =>
	shadowOf(el).querySelector(`button.${cls}`) as HTMLButtonElement;
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
const settle = () => new Promise((resolve) => setTimeout(resolve, 0));

describe('message dialogs', () => {
	it('confirms through the confirm button and exposes alertdialog semantics', async () => {
		const result = confirmDialog({ title: 'Discard?', message: 'Lose changes.\n\nSure?' });
		const [dialog] = dialogs();
		expect(dialog).toBeDefined();
		const box = shadowOf(dialog!).querySelector('.box') as HTMLElement;
		expect(box.getAttribute('role')).toBe('alertdialog');
		expect(box.getAttribute('aria-modal')).toBe('true');
		const title = shadowOf(dialog!).getElementById(box.getAttribute('aria-labelledby')!);
		const body = shadowOf(dialog!).getElementById(box.getAttribute('aria-describedby')!);
		expect(title?.textContent).toBe('Discard?');
		expect(body?.textContent).toBe('Lose changes.\n\nSure?');
		button(dialog!, 'confirm').click();
		expect(await result).toBe(true);
		expect(dialogs()).toHaveLength(0);
	});

	it('uses translated labels and cancels through the cancel button', async () => {
		const result = confirmDialog({
			title: 't',
			message: 'm',
			confirmLabel: 'Ja',
			cancelLabel: 'Nee',
		});
		const [dialog] = dialogs();
		expect(button(dialog!, 'confirm').textContent).toBe('Ja');
		expect(button(dialog!, 'cancel').textContent).toBe('Nee');
		button(dialog!, 'cancel').click();
		expect(await result).toBe(false);
	});

	it('cancels on Escape and the backdrop, confirms on Enter', async () => {
		const escape = confirmDialog({ title: 't', message: 'm' });
		key(dialogs()[0]!, 'Escape');
		expect(await escape).toBe(false);

		const enter = confirmDialog({ title: 't', message: 'm' });
		key(dialogs()[0]!, 'Enter');
		expect(await enter).toBe(true);

		const backdrop = confirmDialog({ title: 't', message: 'm' });
		(shadowOf(dialogs()[0]!).querySelector('.backdrop') as HTMLElement).click();
		expect(await backdrop).toBe(false);
	});

	it('focuses the confirm button, or Cancel for the danger tone, and restores focus', async () => {
		document.body.innerHTML = '<button id="opener">open</button>';
		const opener = document.getElementById('opener') as HTMLButtonElement;
		opener.focus();
		const result = confirmDialog({ title: 't', message: 'm', tone: 'danger' });
		await settle();
		const [dialog] = dialogs();
		expect(shadowOf(dialog!).activeElement).toBe(button(dialog!, 'cancel'));
		expect(button(dialog!, 'confirm').classList.contains('danger')).toBe(true);
		key(dialog!, 'Escape');
		await result;
		expect(document.activeElement).toBe(opener);

		const normal = confirmDialog({ title: 't', message: 'm' });
		await settle();
		expect(shadowOf(dialogs()[0]!).activeElement).toBe(button(dialogs()[0]!, 'confirm'));
		key(dialogs()[0]!, 'Escape');
		await normal;
	});

	it('traps Tab inside the dialog', async () => {
		const result = confirmDialog({ title: 't', message: 'm' });
		await settle();
		const [dialog] = dialogs();
		const cancel = button(dialog!, 'cancel');
		const confirm = button(dialog!, 'confirm');
		confirm.focus();
		const forward = new KeyboardEvent('keydown', {
			key: 'Tab',
			bubbles: true,
			composed: true,
			cancelable: true,
		});
		dialog!.dispatchEvent(forward);
		expect(forward.defaultPrevented).toBe(true);
		expect(shadowOf(dialog!).activeElement).toBe(cancel);
		const back = new KeyboardEvent('keydown', {
			key: 'Tab',
			shiftKey: true,
			bubbles: true,
			composed: true,
			cancelable: true,
		});
		dialog!.dispatchEvent(back);
		expect(back.defaultPrevented).toBe(true);
		expect(shadowOf(dialog!).activeElement).toBe(confirm);
		key(dialog!, 'Escape');
		await result;
	});

	it('prompts for text with a default value, and resolves null on cancel', async () => {
		const result = promptDialog({
			title: 'Rename',
			message: 'New name',
			label: 'Name',
			defaultValue: 'Old',
		});
		await settle();
		const [dialog] = dialogs();
		const input = shadowOf(dialog!).querySelector('input') as HTMLInputElement;
		expect(input.value).toBe('Old');
		expect(input.getAttribute('aria-label')).toBe('Name');
		expect(shadowOf(dialog!).activeElement).toBe(input);
		input.value = 'New';
		key(input, 'Enter');
		expect(await result).toBe('New');

		const cancelled = promptDialog({ title: 'Rename', message: 'x' });
		key(dialogs()[0]!, 'Escape');
		expect(await cancelled).toBeNull();
	});

	it('shows an alert with a single button', async () => {
		const result = alertDialog({ title: 'Oops', message: 'Failed' });
		const [dialog] = dialogs();
		expect(shadowOf(dialog!).querySelector('button.cancel')).toBeNull();
		button(dialog!, 'confirm').click();
		await expect(result).resolves.toBeUndefined();
	});

	it('queues concurrent calls in order and resolves each exactly once', async () => {
		const first = confirmDialog({ title: 'one', message: 'm' });
		const second = promptDialog({ title: 'two', message: 'm', defaultValue: 'x' });
		expect(dialogs()).toHaveLength(1);
		expect(shadowOf(dialogs()[0]!).querySelector('h2')?.textContent).toBe('one');
		const dialog = dialogs()[0]!;
		button(dialog, 'confirm').click();
		button(dialog, 'confirm').click();
		expect(await first).toBe(true);
		await settle();
		expect(dialogs()).toHaveLength(1);
		expect(shadowOf(dialogs()[0]!).querySelector('h2')?.textContent).toBe('two');
		button(dialogs()[0]!, 'confirm').click();
		expect(await second).toBe('x');
		expect(dialogs()).toHaveLength(0);
	});

	it('cancels when the dialog is removed from the page', async () => {
		const result = confirmDialog({ title: 't', message: 'm' });
		dialogs()[0]!.remove();
		expect(await result).toBe(false);
	});

	it('mounts into the host shadow root or the nearest themed ancestor', async () => {
		const outer = document.createElement('div');
		outer.setAttribute('data-office-theme', 'dark');
		const inner = document.createElement('div');
		outer.append(inner);
		const hostEl = document.createElement('div');
		const shadow = hostEl.attachShadow({ mode: 'open' });
		const inShadow = document.createElement('button');
		shadow.append(inShadow);
		document.body.append(outer, hostEl);

		const themed = confirmDialog({ title: 't', message: 'm', host: inner });
		expect(dialogs(outer)).toHaveLength(1);
		key(dialogs(outer)[0]!, 'Escape');
		await themed;

		const shadowed = confirmDialog({ title: 't', message: 'm', host: inShadow });
		expect(dialogs(shadow)).toHaveLength(1);
		expect(dialogs()).toHaveLength(0);
		key(dialogs(shadow)[0]!, 'Escape');
		await shadowed;
		expect(dialogs(shadow)).toHaveLength(0);
	});
});
