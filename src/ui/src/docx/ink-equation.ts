// Word's Ink to Math (Draw tab) and Ink Equation (Equation tab): both open the shared Ink Equation
// dialog with the host's recognizer, as Word opens its Math Input Control from either, and insert
// the equation at the caret (after the selected equation, never over it). Loaded on first use.
import type { EditorView } from 'prosemirror-view';
import { inkMathRecognizer } from '../ink/math-recognizer';
import { focusView } from './focus-view';

export async function openWordInkEquation(options: {
	/** The editor's shadow root (or element), where the dialog goes. */
	host: HTMLElement | ShadowRoot;
	view(): EditorView | undefined;
	locale(): string;
	heading: 'Ink Equation' | 'Ink to Math';
}): Promise<void> {
	const recognizer = inkMathRecognizer();
	if (!recognizer || !options.view()?.editable) return;
	const [{ openInkEquationDialog }, { insertEquationAfterSelection }] = await Promise.all([
		import('../ink/ink-equation-dialog'),
		import('./equation-commands'),
	]);
	openInkEquationDialog({
		host: options.host,
		recognizer,
		locale: options.locale(),
		heading: options.heading,
		insert: (draft) => {
			const view = options.view();
			return Boolean(view && insertEquationAfterSelection(view, draft.latex));
		},
		closed: () => focusView(options.view()),
	});
}
