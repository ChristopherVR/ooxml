import {
	defineMessageDialog,
	OfficeUiMessageDialog,
	type OfficeMessageDialogKind,
	type OfficeMessageDialogResult,
	type OfficeMessageDialogTone,
} from './message-dialog';

/**
 * Where to mount a dialog so it inherits the host's tokens: an element (the dialog goes into its own
 * shadow root when it has one, else to its nearest `[data-office-theme]` ancestor, else to the shadow
 * root or document it lives in), a
 * shadow root, or a document. Omit it to use the page's `document.body`.
 */
export type MessageDialogHost = Node | null | undefined;

export interface MessageDialogOptions {
	/** Translated heading; also the dialog's accessible name. */
	title: string;
	/** Translated body text; newlines are kept. */
	message: string;
	/** Translated name of the confirm button (English "OK" by default). */
	confirmLabel?: string;
	/** Translated name of the Cancel button (English "Cancel" by default). */
	cancelLabel?: string;
	/** `danger` styles the confirm button as destructive and focuses Cancel first. */
	tone?: OfficeMessageDialogTone;
	/** Where to mount; see {@link MessageDialogHost}. */
	host?: MessageDialogHost;
}

export interface PromptDialogOptions extends MessageDialogOptions {
	/** Translated caption above the input (also its accessible name). */
	label?: string;
	defaultValue?: string;
	placeholder?: string;
}

interface Pending {
	options: MessageDialogOptions & Partial<PromptDialogOptions>;
	kind: OfficeMessageDialogKind;
	settle(result: OfficeMessageDialogResult): void;
}

const queue: Pending[] = [];
let active = false;

function mountPoint(host: MessageDialogHost): Node | null {
	const doc = typeof document === 'undefined' ? undefined : document;
	if (!host) return doc?.body ?? null;
	if (host instanceof ShadowRoot) return host;
	if (host instanceof Document) return host.body;
	if (host instanceof Element) {
		// An app that renders into its own shadow root hosts the dialog there, so it shares the
		// app's stacking and tokens (light children of such an element would not be rendered).
		if (host.shadowRoot) return host.shadowRoot;
		const themed = host.closest('[data-office-theme]');
		if (themed) return themed.shadowRoot ?? themed;
		const root = host.getRootNode();
		return root instanceof ShadowRoot ? root : (host.ownerDocument.body ?? null);
	}
	return doc?.body ?? null;
}

function run(pending: Pending): void {
	const { options, kind } = pending;
	const parent = mountPoint(options.host);
	if (!parent) {
		pending.settle({ confirmed: false, value: '' });
		return;
	}
	defineMessageDialog();
	const dialog = parent.ownerDocument?.createElement('office-ui-message-dialog') ?? null;
	if (!(dialog instanceof OfficeUiMessageDialog)) {
		pending.settle({ confirmed: false, value: '' });
		return;
	}
	dialog.kind = kind;
	dialog.tone = options.tone ?? 'default';
	dialog.heading = options.title;
	dialog.message = options.message;
	if (options.confirmLabel) dialog.confirmLabel = options.confirmLabel;
	if (options.cancelLabel) dialog.cancelLabel = options.cancelLabel;
	if (options.label) dialog.inputLabel = options.label;
	dialog.value = options.defaultValue ?? '';
	dialog.placeholder = options.placeholder ?? '';
	const { resultEvent } = OfficeUiMessageDialog;
	const onResult = (event: Event): void => {
		event.stopPropagation();
		dialog.removeEventListener(resultEvent, onResult);
		const detail = (event as CustomEvent<OfficeMessageDialogResult>).detail;
		// Remove first so focus returns to the invoker before the caller continues.
		dialog.remove();
		pending.settle(detail);
	};
	dialog.addEventListener(resultEvent, onResult);
	parent.appendChild(dialog);
}

function next(): void {
	if (active) return;
	const pending = queue.shift();
	if (!pending) return;
	active = true;
	let settled = false;
	const original = pending.settle;
	pending.settle = (result) => {
		if (settled) return;
		settled = true;
		active = false;
		original(result);
		next();
	};
	run(pending);
}

function ask(
	kind: OfficeMessageDialogKind,
	options: MessageDialogOptions & Partial<PromptDialogOptions>,
): Promise<OfficeMessageDialogResult> {
	return new Promise((resolve) => {
		queue.push({ kind, options, settle: resolve });
		next();
	});
}

/**
 * Ask a yes/no question in an in-app dialog. Resolves `true` for the confirm button or Enter and
 * `false` for Cancel, Escape, the backdrop, or when there is no DOM. Calls made while a dialog is
 * open are queued and shown one after another in call order, so every promise resolves exactly once.
 */
export async function confirmDialog(options: MessageDialogOptions): Promise<boolean> {
	return (await ask('confirm', options)).confirmed;
}

/** Ask for a line of text. Resolves the text on confirm (not trimmed) or `null` when cancelled. */
export async function promptDialog(options: PromptDialogOptions): Promise<string | null> {
	const result = await ask('prompt', options);
	return result.confirmed ? result.value : null;
}

/** Tell the user something; resolves when it is dismissed (button, Enter, Escape or backdrop). */
export async function alertDialog(options: MessageDialogOptions): Promise<void> {
	await ask('alert', options);
}
