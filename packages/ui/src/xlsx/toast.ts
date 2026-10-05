/** Short notices in the editor's corner (save compatibility, command failures). */
export type ToastKind = 'info' | 'warning' | 'error';

export interface Toaster {
	readonly element: HTMLElement;
	show(message: string, kind?: ToastKind, timeoutMs?: number): void;
	clear(): void;
}

const MAX_TOASTS = 3;

export function createToaster(doc: Document, closeLabel: () => string): Toaster {
	const element = doc.createElement('div');
	element.className = 'xve-toasts';
	element.setAttribute('aria-live', 'polite');
	return {
		element,
		show(message, kind = 'info', timeoutMs = kind === 'error' ? 10_000 : 6000) {
			const toast = doc.createElement('div');
			toast.className = 'xve-toast';
			toast.dataset.kind = kind;
			toast.setAttribute('role', kind === 'error' ? 'alert' : 'status');
			const text = doc.createElement('span');
			text.className = 'xve-toast-text';
			text.textContent = message;
			const close = doc.createElement('button');
			close.type = 'button';
			close.className = 'xve-icon-button';
			close.textContent = 'x';
			close.setAttribute('aria-label', closeLabel());
			close.addEventListener('click', () => toast.remove());
			toast.append(text, close);
			element.append(toast);
			while (element.childElementCount > MAX_TOASTS) element.firstElementChild?.remove();
			setTimeout(() => toast.remove(), timeoutMs);
		},
		clear() {
			element.replaceChildren();
		},
	};
}
