import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { OFFICE_TOAST_VISIBLE_LIMIT, type OfficeToast } from './controls.js';
import { registerOfficeUi } from './index.js';

// Adapted from pptx-viewer's control-primitives and chrome-controls tests (the elements moved here).
beforeAll(() => registerOfficeUi());
afterEach(() => document.body.replaceChildren());

type Stateful<T> = HTMLElement & { state: T };
const q = <T extends Element>(host: HTMLElement, selector: string) =>
	host.shadowRoot!.querySelector<T>(selector);
const listen = (host: HTMLElement, type: string) => {
	const spy = vi.fn();
	host.addEventListener(type, (event) => spy((event as CustomEvent).detail));
	return spy;
};
const key = (el: HTMLElement, name: string) =>
	el.dispatchEvent(new KeyboardEvent('keydown', { key: name, bubbles: true, cancelable: true }));

describe('office-ui-search', () => {
	function make(attributes: Record<string, string> = {}) {
		const search = document.createElement('office-ui-search') as HTMLElement & {
			value: string;
			disabled: boolean;
			placeholder: string;
		};
		for (const [name, value] of Object.entries(attributes)) search.setAttribute(name, value);
		document.body.append(search);
		return { search, input: search.shadowRoot!.querySelector('input')! };
	}

	it('names the input from aria-label, falling back to the placeholder', () => {
		expect(make({ placeholder: 'Search recent' }).input.getAttribute('aria-label')).toBe(
			'Search recent',
		);
		expect(make({ placeholder: 'x', 'aria-label': 'Find' }).input.getAttribute('aria-label')).toBe(
			'Find',
		);
		const { search, input } = make();
		search.placeholder = 'Search';
		expect(input.placeholder).toBe('Search');
	});

	it('forwards one input event per keystroke, keeps value in sync and disables', () => {
		const { search, input } = make();
		const onInput = vi.fn();
		search.addEventListener('input', onInput);
		input.value = 'a';
		input.dispatchEvent(new Event('input', { bubbles: true }));
		input.value = 'ab';
		input.dispatchEvent(new Event('input', { bubbles: true }));
		expect(onInput).toHaveBeenCalledTimes(2);
		expect(search.value).toBe('ab');
		search.value = 'set';
		expect(input.value).toBe('set');
		search.disabled = true;
		expect(input.disabled).toBe(true);
	});
});

describe('office-ui-radio', () => {
	type Radio = HTMLElement & { checked: boolean; disabled: boolean; value: string };
	function group(count = 3, name = 'g', tag = 'office-ui-radio'): Radio[] {
		return Array.from({ length: count }, (_, i) => {
			const radio = document.createElement(tag) as Radio;
			radio.setAttribute('name', name);
			radio.setAttribute('value', String(i));
			document.body.append(radio);
			return radio;
		});
	}

	it('exposes role, checked state and one roving tab stop; checking unchecks peers silently', () => {
		const [a, b, c] = group();
		expect(a!.getAttribute('role')).toBe('radio');
		expect([a, b, c].map((r) => r!.tabIndex)).toStrictEqual([0, -1, -1]);
		const onChange = vi.fn();
		document.body.addEventListener('change', onChange);
		a!.checked = true;
		b!.checked = true;
		expect(a!.checked).toBe(false);
		expect([a, b, c].map((r) => r!.tabIndex)).toStrictEqual([-1, 0, -1]);
		expect(onChange).not.toHaveBeenCalled();
	});

	it('click and space select once and emit input then change; names stay independent', () => {
		const [a, b] = group();
		const [x] = group(2, 'other');
		const seen: string[] = [];
		b!.addEventListener('input', () => seen.push('input'));
		b!.addEventListener('change', () => seen.push('change'));
		b!.click();
		b!.click();
		expect(seen).toStrictEqual(['input', 'change']);
		x!.checked = true;
		key(a!, ' ');
		expect(a!.checked).toBe(true);
		expect(b!.checked).toBe(false);
		expect(x!.checked).toBe(true);
	});

	it('arrows, Home and End move and select, wrapping and skipping disabled radios', () => {
		const [a, b, c] = group();
		a!.checked = true;
		b!.disabled = true;
		a!.focus();
		key(a!, 'ArrowDown');
		expect(c!.checked).toBe(true);
		expect(document.activeElement).toBe(c);
		key(c!, 'ArrowRight');
		expect(a!.checked).toBe(true);
		key(a!, 'End');
		expect(c!.checked).toBe(true);
		expect(b!.getAttribute('aria-disabled')).toBe('true');
		expect(b!.tabIndex).toBe(-1);
	});

	it('groups only radios of the same tag, so product aliases form their own groups', () => {
		class Alias extends (customElements.get('office-ui-radio') as CustomElementConstructor) {}
		if (!customElements.get('test-ui-radio')) customElements.define('test-ui-radio', Alias);
		const [office] = group(1, 'shared');
		const [alias] = group(1, 'shared', 'test-ui-radio');
		office!.checked = true;
		alias!.checked = true;
		expect(office!.checked).toBe(true);
		expect(alias!.checked).toBe(true);
	});
});

describe('office-ui-ribbon-toggle', () => {
	it('reflects state and emits the command with the new checked value', () => {
		const toggle = document.createElement('office-ui-ribbon-toggle');
		toggle.setAttribute('label', 'Ruler');
		toggle.setAttribute('command', 'ruler');
		document.body.append(toggle);
		const box = q<HTMLElement & { checked: boolean }>(toggle, 'office-ui-checkbox')!;
		expect(box.getAttribute('aria-label')).toBe('Ruler');
		const spy = listen(toggle, 'office-toggle');
		q<HTMLLabelElement>(toggle, 'label')!.click();
		expect(spy).toHaveBeenCalledWith({ command: 'ruler', checked: true });
		toggle.setAttribute('disabled', '');
		expect((box as unknown as { disabled: boolean }).disabled).toBe(true);
	});
});

describe('office-ui-read-only-banner', () => {
	type Banner = Stateful<{
		kind: string | null;
		message: string;
		passwordPromptOpen?: boolean;
		passwordError?: string | null;
		checkingPassword?: boolean;
	}>;
	const mount = () => {
		const host = document.createElement('office-ui-read-only-banner') as Banner;
		host.state = { kind: 'markAsFinal', message: 'The author marked this final.' };
		document.body.append(host);
		return host;
	};

	it('renders the message and hooks and emits one intent per action', () => {
		const host = mount();
		expect(host.dataset.testid).toBe('office-readonly-banner');
		expect(host.dataset.kind).toBe('markAsFinal');
		expect(q(host, '.text')!.textContent).toBe(
			'Read-only recommended: The author marked this final.',
		);
		const spy = listen(host, 'office-read-only-request');
		q<HTMLButtonElement>(host, '[data-testid="office-readonly-edit-anyway"]')!.click();
		q<HTMLButtonElement>(host, '[data-testid="office-readonly-dismiss"]')!.click();
		expect(spy.mock.calls).toStrictEqual([[{ id: 'editAnyway' }], [{ id: 'dismiss' }]]);
	});

	it('opens the password form with focus, submits, reports errors and clears on close', () => {
		const host = mount();
		host.state = { ...host.state, passwordPromptOpen: true };
		const input = q<HTMLInputElement>(host, '[data-testid="office-readonly-password-input"]')!;
		expect(host.shadowRoot!.activeElement).toBe(input);
		const spy = listen(host, 'office-read-only-request');
		input.value = 'letmeedit123';
		q<HTMLButtonElement>(host, '[data-testid="office-readonly-unlock"]')!.form!.requestSubmit();
		expect(spy).toHaveBeenCalledWith({ id: 'submitPassword', password: 'letmeedit123' });
		host.state = { ...host.state, passwordError: 'Wrong password', checkingPassword: true };
		const error = q<HTMLElement>(host, '[data-testid="office-readonly-password-error"]')!;
		expect(error.textContent).toBe('Wrong password');
		expect(input.getAttribute('aria-invalid')).toBe('true');
		expect(input.disabled).toBe(true);
		host.state = { ...host.state, passwordPromptOpen: false, passwordError: null };
		expect(input.value).toBe('');
	});
});

describe('office-ui-paste-options', () => {
	const OPTIONS = [
		{ id: 'keep-source', label: 'Keep Source Formatting' },
		{ id: 'picture', label: 'Picture' },
	];

	it('positions the strip, names it and emits the chosen format then dismiss', () => {
		const host = document.createElement('office-ui-paste-options') as Stateful<unknown>;
		host.state = { left: 100, top: 50, options: OPTIONS, label: 'Paste Options' };
		document.body.append(host);
		expect(host.style.left).toBe('calc(100px + var(--office-space-1, 4px))');
		expect(host.style.top).toBe('calc(50px + var(--office-space-1, 4px))');
		expect(q(host, '[role="toolbar"]')!.getAttribute('aria-label')).toBe('Paste Options');
		const buttons = host.shadowRoot!.querySelectorAll('button');
		expect(buttons[1]!.textContent).toBe('Picture');
		const chosen = listen(host, 'office-paste-options-request');
		const dismissed = listen(host, 'office-paste-options-dismiss');
		buttons[1]!.click();
		expect(chosen).toHaveBeenCalledWith({ format: 'picture' });
		expect(dismissed).toHaveBeenCalledOnce();
	});

	it('arms outside dismissal after a task and ignores presses on the strip', async () => {
		vi.useFakeTimers();
		try {
			const host = document.createElement('office-ui-paste-options') as Stateful<unknown>;
			host.state = { left: 0, top: 0, options: OPTIONS };
			document.body.append(host);
			const spy = listen(host, 'office-paste-options-dismiss');
			document.body.dispatchEvent(new Event('pointerdown', { bubbles: true }));
			expect(spy).not.toHaveBeenCalled();
			await vi.runAllTimersAsync();
			q<HTMLElement>(host, '[role="toolbar"]')!.dispatchEvent(
				new Event('pointerdown', { bubbles: true, composed: true }),
			);
			expect(spy).not.toHaveBeenCalled();
			document.body.dispatchEvent(new Event('pointerdown', { bubbles: true }));
			expect(spy).toHaveBeenCalledOnce();
			host.remove();
			document.body.dispatchEvent(new Event('pointerdown', { bubbles: true }));
			expect(spy).toHaveBeenCalledOnce();
		} finally {
			vi.useRealTimers();
		}
	});
});

describe('office-ui-toasts', () => {
	const toast = (n: number, severity: 'info' | 'warning' = 'warning'): OfficeToast => ({
		id: `id${n}`,
		code: `CODE_${n}`,
		severity,
		message: `Message ${n}`,
	});

	it('hides when empty, caps the list and shows the overflow count', () => {
		const host = document.createElement('office-ui-toasts') as Stateful<unknown>;
		host.state = { toasts: [] };
		document.body.append(host);
		expect(host.hidden).toBe(true);
		host.state = {
			toasts: Array.from({ length: OFFICE_TOAST_VISIBLE_LIMIT + 2 }, (_, i) => toast(i)),
			overflowCount: 1,
		};
		const items = host.shadowRoot!.querySelectorAll('[data-testid="office-toast"]');
		expect(items).toHaveLength(OFFICE_TOAST_VISIBLE_LIMIT);
		expect(items[0]!.getAttribute('data-code')).toBe('CODE_0');
		expect(q(host, '.overflow')!.textContent).toBe('+3');
	});

	it('emits dismiss and dismiss-all intents and keeps focus when nothing visible changed', () => {
		const host = document.createElement('office-ui-toasts') as Stateful<{ toasts: OfficeToast[] }>;
		host.state = { toasts: [toast(1), toast(2, 'info')] };
		document.body.append(host);
		const spy = listen(host, 'office-toasts-request');
		const dismiss = host.shadowRoot!.querySelectorAll<HTMLButtonElement>(
			'[data-testid="office-toast-dismiss"]',
		);
		expect(dismiss[0]!.getAttribute('aria-label')).toBe('Dismiss');
		dismiss[1]!.click();
		q<HTMLButtonElement>(host, '[data-testid="office-toasts-dismiss-all"]')!.click();
		expect(spy.mock.calls).toStrictEqual([
			[{ id: 'dismiss', toastId: 'id2' }],
			[{ id: 'dismissAll' }],
		]);
		dismiss[1]!.focus();
		host.state = { ...host.state };
		expect(host.shadowRoot!.activeElement).toBe(dismiss[1]);
	});
});

describe('office-ui-dialog-footer', () => {
	it('renders variants, icons, hooks and busy state, emits ids and keeps focus', () => {
		const host = document.createElement('office-ui-dialog-footer') as Stateful<unknown>;
		host.state = {
			actions: [
				{ id: 'remove', label: 'Remove link', align: 'start', testId: 'remove-hook' },
				{ id: 'ok', label: 'Print', variant: 'primary', icon: 'print' },
				{ id: 'wait', label: 'Wait', busy: true },
			],
		};
		document.body.append(host);
		const [remove, ok, wait] = Array.from(host.shadowRoot!.querySelectorAll('button'));
		expect(remove!.dataset.testid).toBe('remove-hook');
		expect(remove!.classList.contains('start')).toBe(true);
		expect(ok!.className).toBe('primary');
		expect(ok!.querySelector('svg path')).not.toBeNull();
		expect(wait!.disabled).toBe(true);
		expect(wait!.getAttribute('aria-busy')).toBe('true');
		const spy = listen(host, 'office-dialog-footer-request');
		ok!.click();
		expect(spy).toHaveBeenCalledWith({ id: 'ok' });
		ok!.focus();
		host.state = {
			actions: [
				{ id: 'ok', label: 'Print' },
				{ id: 'new', label: 'New' },
			],
		};
		expect((host.shadowRoot!.activeElement as HTMLElement).dataset.action).toBe('ok');
	});

	it('lets a product subclass keep its own event name', () => {
		const Base = customElements.get('office-ui-dialog-footer') as CustomElementConstructor & {
			requestEvent: string;
		};
		class Legacy extends Base {
			static override requestEvent = 'legacy-footer-request';
		}
		if (!customElements.get('legacy-dialog-footer'))
			customElements.define('legacy-dialog-footer', Legacy);
		const host = document.createElement('legacy-dialog-footer') as Stateful<unknown>;
		host.state = { actions: [{ id: 'cancel', label: 'Cancel' }] };
		document.body.append(host);
		const spy = listen(host, 'legacy-footer-request');
		host.shadowRoot!.querySelector('button')!.click();
		expect(spy).toHaveBeenCalledWith({ id: 'cancel' });
	});
});
