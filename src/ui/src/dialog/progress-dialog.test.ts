import { registerOfficeUi } from '../index';
import type { OfficeProgressDialogState } from './progress-dialog';

beforeAll(() => registerOfficeUi());
afterEach(() => document.body.replaceChildren());

type ProgressDialog = HTMLElement & { state: OfficeProgressDialogState };

const running: OfficeProgressDialogState = {
	open: true,
	title: 'Export as PDF',
	progress: 42.4,
	status: 'Rendering slide 3 of 7...',
	cancelLabel: 'Stop',
};

function mount(state: OfficeProgressDialogState = running) {
	document.body.innerHTML = '<button id="opener">Export</button>';
	const opener = document.getElementById('opener') as HTMLButtonElement;
	opener.focus();
	const dialog = document.createElement('office-ui-progress-dialog') as ProgressDialog;
	document.body.append(dialog);
	dialog.state = state;
	const root = dialog.shadowRoot!;
	return { dialog, opener, root, cancel: () => root.querySelector('button')! };
}

const key = (el: Element, k: string) => {
	const event = new KeyboardEvent('keydown', {
		key: k,
		bubbles: true,
		composed: true,
		cancelable: true,
	});
	el.dispatchEvent(event);
	return event;
};

describe('office-ui-progress-dialog', () => {
	it('is a modal dialog named by its heading, with a determinate progress bar', () => {
		const { dialog, root } = mount();
		expect(dialog.hasAttribute('open')).toBeTruthy();
		const box = root.querySelector('[role="dialog"]')!;
		expect(box.getAttribute('aria-modal')).toBe('true');
		expect(root.getElementById(box.getAttribute('aria-labelledby')!)!.textContent).toBe(
			'Export as PDF',
		);
		const bar = root.querySelector('[role="progressbar"]')!;
		expect(bar.getAttribute('aria-labelledby')).toBe('title');
		expect([
			bar.getAttribute('aria-valuemin'),
			bar.getAttribute('aria-valuemax'),
			bar.getAttribute('aria-valuenow'),
		]).toStrictEqual(['0', '100', '42']);
		expect(root.querySelector<HTMLElement>('.fill')!.style.width).toBe('42%');
		const status = root.getElementById(box.getAttribute('aria-describedby')!)!;
		expect(status.textContent).toBe('Rendering slide 3 of 7...');
		expect(status.getAttribute('aria-live')).toBe('polite');
		expect(root.querySelector('.percent')!.textContent).toBe('42%');
	});

	it('clamps the progress into 0..100', () => {
		const { dialog, root } = mount({ ...running, progress: 140 });
		const bar = () => root.querySelector('[role="progressbar"]')!;
		expect(bar().getAttribute('aria-valuenow')).toBe('100');
		dialog.state = { ...running, progress: Number.NaN };
		expect(bar().getAttribute('aria-valuenow')).toBe('0');
		dialog.state = { ...running, progress: -5 };
		expect(bar().getAttribute('aria-valuenow')).toBe('0');
	});

	it('moves focus to Cancel, keeps it there and returns it when closed', () => {
		const { dialog, opener, root, cancel } = mount();
		expect(root.activeElement).toBe(cancel());
		expect(key(cancel(), 'Tab').defaultPrevented).toBeTruthy();
		expect(root.activeElement).toBe(cancel());
		dialog.state = { ...running, open: false };
		expect(dialog.hasAttribute('open')).toBeFalsy();
		expect(root.querySelector('[role="dialog"]')).toBeNull();
		expect(document.activeElement).toBe(opener);
	});

	it('returns focus to an opener inside a shadow root, not to its host', () => {
		document.body.innerHTML = '<div id="host"></div>';
		const host = document.getElementById('host')!;
		const inner = host.attachShadow({ mode: 'open' });
		inner.innerHTML = '<button>Export</button>';
		const opener = inner.querySelector('button')!;
		opener.focus();
		const dialog = document.createElement('office-ui-progress-dialog') as ProgressDialog;
		document.body.append(dialog);
		dialog.state = running;
		expect(dialog.shadowRoot!.activeElement).toBe(dialog.shadowRoot!.querySelector('button'));
		dialog.state = { ...running, open: false };
		expect(inner.activeElement).toBe(opener);
	});

	it('keeps focus on Cancel when the backdrop or the box is pressed', () => {
		const { root, cancel } = mount();
		const press = (el: Element) => {
			const event = new MouseEvent('mousedown', {
				bubbles: true,
				composed: true,
				cancelable: true,
			});
			el.dispatchEvent(event);
			return event.defaultPrevented;
		};
		// Pressing a part that cannot take focus would move it to the page behind the overlay.
		expect(press(root.querySelector('.backdrop')!)).toBeTruthy();
		expect(press(root.querySelector('h2')!)).toBeTruthy();
		expect(press(cancel())).toBeFalsy();
	});

	it('emits a cancel request from Cancel only: Escape does not end the job', () => {
		const { dialog, cancel } = mount();
		const cancelled = vi.fn();
		dialog.addEventListener('office-progress-cancel', cancelled);
		expect(key(cancel(), 'Escape').defaultPrevented).toBeTruthy();
		expect(cancelled).not.toHaveBeenCalled();
		cancel().click();
		expect(cancelled).toHaveBeenCalledOnce();
		expect(cancel().textContent).toBe('Stop');
		// The host closes it once the job has stopped.
		expect(dialog.hasAttribute('open')).toBeTruthy();
	});

	it('renders nothing while closed, and Cancel reads in English by default', () => {
		const { dialog, root } = mount({ ...running, open: false });
		expect(root.querySelector('[role="dialog"]')).toBeNull();
		const { cancelLabel: _, ...unlabelled } = running;
		dialog.state = unlabelled;
		expect(root.querySelector('button')!.textContent).toBe('Cancel');
	});
});
