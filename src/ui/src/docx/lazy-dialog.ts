// A dialog built on first open from a module loaded then, so an editor that never opens it never
// downloads it (the chart dialogs bring the chart picker, the data sheet and the ChartEx options).
import type { FormatDialog } from './font-dialog';
import type { EditorLocale } from './localization';

export function lazyDialog(load: () => Promise<FormatDialog>): FormatDialog {
	// Its place among the dialogs; the real dialog's element goes inside once it is built.
	const element = document.createElement('div');
	element.style.display = 'contents';
	let real: FormatDialog | undefined;
	let loading: Promise<FormatDialog> | undefined;
	let locale: EditorLocale | undefined;
	const get = () =>
		(loading ??= load().then(
			(dialog) => {
				real = dialog;
				element.append(dialog.element);
				if (locale) dialog.setLocale(locale);
				return dialog;
			},
			(cause: unknown) => {
				loading = undefined;
				throw cause;
			},
		));
	return {
		element,
		open: (opener) => void get().then((dialog) => dialog.open(opener)),
		close: () => real?.close(),
		setLocale(value) {
			locale = value;
			real?.setLocale(value);
		},
		get isOpen() {
			return real?.isOpen ?? false;
		},
	};
}
