// A text field that colours the references of a formula, used by the in-cell editor and the
// formula bar. A transparent <textarea> (which owns the caret, IME and native undo) sits over a
// backdrop that paints the same text with coloured reference spans.
import { coloredReferences, isFormulaText, type ColoredReference } from './formula-text.js';
import { h } from './dom.js';

export interface FormulaField {
	readonly root: HTMLElement;
	readonly input: HTMLTextAreaElement;
	/** Sets the text (and caret) without firing input events. */
	setText(text: string, caret?: number): void;
	/** Repaints the backdrop from the current input value. */
	refresh(): ColoredReference[];
	references(): ColoredReference[];
}

export function createFormulaField(doc: Document, className: string, label: string): FormulaField {
	const root = h(doc, 'div', `xg-field ${className}`);
	const backdrop = h(doc, 'div', 'xg-field-backdrop', { 'aria-hidden': 'true' });
	const input = h(doc, 'textarea', 'xg-field-input', {
		spellcheck: 'false',
		autocomplete: 'off',
		autocapitalize: 'off',
		'aria-label': label,
		rows: '1',
	});
	input.wrap = 'soft';
	root.append(backdrop, input);
	let refs: ColoredReference[] = [];

	const refresh = (): ColoredReference[] => {
		const text = input.value;
		const formula = isFormulaText(text);
		root.classList.toggle('xg-formula', formula);
		refs = formula ? coloredReferences(text) : [];
		backdrop.replaceChildren();
		if (formula) {
			let at = 0;
			for (const ref of refs) {
				if (ref.start > at) backdrop.append(doc.createTextNode(text.slice(at, ref.start)));
				const span = h(doc, 'span', 'xg-ref');
				span.style.color = ref.color;
				span.textContent = text.slice(ref.start, ref.end);
				backdrop.append(span);
				at = ref.end;
			}
			// A trailing newline needs a character after it to take up a line in the backdrop.
			backdrop.append(doc.createTextNode(`${text.slice(at)}​`));
		}
		backdrop.scrollTop = input.scrollTop;
		backdrop.scrollLeft = input.scrollLeft;
		return refs;
	};
	input.addEventListener('input', () => refresh());
	input.addEventListener('scroll', () => {
		backdrop.scrollTop = input.scrollTop;
		backdrop.scrollLeft = input.scrollLeft;
	});

	return {
		root,
		input,
		setText(text, caret = text.length) {
			if (input.value !== text) input.value = text;
			try {
				input.setSelectionRange(caret, caret);
			} catch {
				// Detached or hidden inputs may refuse a selection; the caret is applied on focus.
			}
			refresh();
		},
		refresh,
		references: () => refs,
	};
}
