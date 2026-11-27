// A dialog built on first open from a module loaded then, so an editor that never opens it never
// downloads it (the chart dialogs bring the chart picker, the data sheet and the ChartEx options).
// While it loads it counts as open: closing it (Escape, Close All, another dialog) cancels the
// open, a second open is ignored, and a load that fails is reported instead of opening later.
import type { FormatDialog } from './font-dialog';
import type { EditorLocale } from './localization';

export function lazyDialog(
	load: () => Promise<FormatDialog>,
	failed?: (error: Error) => void,
): FormatDialog {
	// Its place among the dialogs; the real dialog's element goes inside once it is built.
	const element = document.createElement('div');
	element.style.display = 'contents';
	let real: FormatDialog | undefined;
	let loading: Promise<FormatDialog> | undefined;
	let locale: EditorLocale | undefined;
	/** The open waiting for the module; cleared by close. */
	let pending: object | undefined;
	const get = () =>
		(loading ??= load().then(
			(dialog) => {
				real = dialog;
				element.append(dialog.element);
				if (locale) dialog.setLocale(locale);
				return dialog;
			},
			(cause: unknown) => {
				// Tried again on the next open.
				loading = undefined;
				throw cause;
			},
		));
	return {
		element,
		open(opener) {
			if (real) {
				real.open(opener);
				return;
			}
			if (pending) return;
			const token = (pending = {});
			get().then(
				(dialog) => {
					if (pending !== token) return;
					pending = undefined;
					dialog.open(opener);
				},
				(cause: unknown) => {
					if (pending !== token) return;
					pending = undefined;
					failed?.(cause instanceof Error ? cause : new Error(String(cause)));
				},
			);
		},
		close() {
			pending = undefined;
			real?.close();
		},
		setLocale(value) {
			locale = value;
			real?.setLocale(value);
		},
		get isOpen() {
			return pending !== undefined || (real?.isOpen ?? false);
		},
	};
}
