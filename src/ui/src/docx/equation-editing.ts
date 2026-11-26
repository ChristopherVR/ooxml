// Editing an equation in place: while an equation is selected, a box under it holds the caret and
// the equation's linear text (Unicode or LaTeX). Typing builds the text up into Office Math as it
// changes, a space after a control word runs math AutoCorrect (`\alpha`), and the Equation tab's
// structures and symbols insert at the caret, selecting the first placeholder. The document shows
// the built-up equation (Professional) or its linear text (Linear). An equation linear text cannot
// rebuild exactly (`editableEquation`) is shown read-only, so editing never loses content.
import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { EQUATION_PLACEHOLDER, type EquationInputFormat } from 'ooxml-core/math';
import type { EditableEquation } from 'ooxml-core/math';
import { equationKit as kit } from './equation-kit';
import { setEquationXml } from './equation-commands';
import { emit } from './events';
import editorCss from './styles/equation-editor.css?raw';
import { focusView } from './focus-view';
import { translate, type EditorLocale, type LocalizationKey } from './localization';
import { normalTextAround, placeholderArgument } from './equation-linear-text';

export interface EquationEditingHost {
	view(): EditorView | undefined;
	toolbar(): HTMLElement | undefined;
	locale(): EditorLocale;
	format(): EquationInputFormat;
	setFormat(format: EquationInputFormat): void;
	refresh(): void;
}

/** The galleries' and editor's styles, which load with them rather than with the editor. */
export function equationStyle(): HTMLStyleElement {
	const style = document.createElement('style');
	style.textContent = editorCss;
	return style;
}

const READ_ONLY: LocalizationKey =
	'This equation has content its linear text cannot hold, so it is kept as it is.';
/** Status messages wait for typing to pause, so a screen reader is not interrupted per key. */
const STATUS_DELAY_MS = 700;

export class EquationEditingBox {
	readonly element: HTMLDivElement;
	readonly input: HTMLInputElement;
	private readonly status: HTMLElement;
	private pos: number | undefined;
	/** The markup this box last wrote or read, so edits from elsewhere (undo, collaboration) reload it. */
	private committed = '';
	/** How the equation is edited, or undefined when it is read-only. */
	private edit: EditableEquation | undefined;
	private statusTimer: ReturnType<typeof setTimeout> | undefined;
	private frame = 0;
	private observer: ResizeObserver | undefined;
	private observed: HTMLElement | undefined;
	private readonly reposition = () => {
		if (!this.frame) this.frame = requestAnimationFrame(() => ((this.frame = 0), this.place()));
	};

	constructor(private readonly host: EquationEditingHost) {
		this.element = document.createElement('div');
		this.element.className = 'dve-equation-editor';
		this.element.setAttribute('role', 'group');
		this.input = document.createElement('input');
		this.input.type = 'text';
		this.input.spellcheck = false;
		this.input.autocomplete = 'off';
		this.input.className = 'dve-equation-editor-input';
		this.status = document.createElement('span');
		this.status.className = 'dve-equation-editor-status';
		this.status.setAttribute('role', 'status');
		this.element.append(equationStyle(), this.input, this.status);
		this.input.addEventListener('input', () => this.commit());
		this.input.addEventListener('keydown', (event) => this.keydown(event));
		this.input.addEventListener('select', () => host.refresh());
		this.input.addEventListener('keyup', () => host.refresh());
		this.input.addEventListener('pointerup', () => host.refresh());
		if (typeof ResizeObserver !== 'undefined') this.observer = new ResizeObserver(this.reposition);
	}

	get attachedAt(): number | undefined {
		return this.pos;
	}

	/** Whether the equation can be edited here; read-only ones keep their markup untouched. */
	get editable(): boolean {
		return this.edit !== undefined && !this.input.readOnly;
	}

	/** Whether the caret is in normal text, for the Normal Text toggle. */
	get normalText(): boolean {
		if (this.pos === undefined) return false;
		const caret = this.input.selectionStart ?? this.input.value.length;
		return Boolean(normalTextAround(this.input.value, caret, this.host.format()));
	}

	/**
	 * Opens the box on the equation at `pos`, or brings it up to date with the document. `same`
	 * says the selection is still on the equation the box is open on (its position may have moved
	 * with edits elsewhere); the text is then reloaded only when the equation itself changed.
	 */
	sync(view: EditorView, pos: number | undefined, focus: boolean, same = pos === this.pos): void {
		if (pos === undefined) return this.detach();
		const node = view.state.doc.nodeAt(pos);
		if (node?.type.name !== 'equation') return this.detach();
		const omml = String(node.attrs.omml);
		const fresh = !same || this.pos === undefined;
		if (fresh || omml !== this.committed) this.load(omml);
		this.pos = pos;
		this.input.readOnly = !view.editable || !this.edit;
		const locale = this.host.locale();
		const label = translate(locale, 'Equation editor');
		this.element.setAttribute('aria-label', label);
		this.input.setAttribute('aria-label', label);
		this.input.placeholder = translate(locale, 'Type equation here.');
		if (!this.element.isConnected) {
			const root = this.host.toolbar()?.getRootNode();
			(root instanceof ShadowRoot ? root : document.body).append(this.element);
			window.addEventListener('scroll', this.reposition, true);
			window.addEventListener('resize', this.reposition);
		}
		this.place();
		this.reposition();
		if (focus && (fresh || !this.input.matches(':focus'))) {
			this.input.focus({ preventScroll: true });
			const placeholder = this.input.value.indexOf(EQUATION_PLACEHOLDER);
			if (placeholder >= 0) this.input.setSelectionRange(placeholder, placeholder + 1);
			else this.input.setSelectionRange(this.input.value.length, this.input.value.length);
		}
	}

	detach(): void {
		this.pos = undefined;
		this.committed = '';
		this.edit = undefined;
		if (this.frame) cancelAnimationFrame(this.frame);
		this.frame = 0;
		clearTimeout(this.statusTimer);
		this.observer?.disconnect();
		this.observed?.classList.remove('dve-equation-editing');
		this.observed = undefined;
		window.removeEventListener('scroll', this.reposition, true);
		window.removeEventListener('resize', this.reposition);
		this.element.remove();
	}

	private setStatus(key: LocalizationKey | undefined, now = false): void {
		clearTimeout(this.statusTimer);
		const text = key ? translate(this.host.locale(), key) : '';
		if (!text || now) this.status.textContent = text;
		else this.statusTimer = setTimeout(() => (this.status.textContent = text), STATUS_DELAY_MS);
	}

	private load(omml: string): void {
		this.edit = kit().editableEquation(omml, this.host.format());
		this.input.value = this.edit?.text ?? kit().linearizeEquation(omml, this.host.format()) ?? '';
		this.committed = omml;
		this.setStatus(this.edit ? undefined : READ_ONLY, true);
	}

	/** The equation on screen: in Draft its typeset maths, in Print Layout the page selection. */
	private anchor(view: EditorView): { rect: DOMRect; node?: HTMLElement } | undefined {
		const dom = this.pos === undefined ? null : view.nodeDOM(this.pos);
		// The equation's node is laid out in Draft (in Print Layout the hidden editing surface is not):
		// until its maths is drawn it has no width yet, and the observer places the box again then.
		if (dom instanceof HTMLElement && dom.getClientRects().length > 0) {
			const math = dom.querySelector('math');
			const rect = (math ?? dom).getBoundingClientRect();
			return { rect: rect.width > 0 ? rect : dom.getBoundingClientRect(), node: dom };
		}
		const root = view.dom.getRootNode() as ParentNode;
		const boxes = [...root.querySelectorAll<HTMLElement>('.dve-print-selection')]
			.map((box) => box.getBoundingClientRect())
			.filter((box) => box.width > 0);
		const rect = boxes.at(-1);
		return rect && { rect };
	}

	/** Puts the box under the equation; in Draft the equation makes room for it below. */
	private place(): void {
		const view = this.host.view();
		if (!view || view.isDestroyed || this.pos === undefined) return;
		const found = this.anchor(view);
		const rect = found?.rect ?? view.dom.getBoundingClientRect();
		if (found?.node !== this.observed) {
			this.observed?.classList.remove('dve-equation-editing');
			this.observer?.disconnect();
			this.observed = found?.node;
			if (this.observed) this.observer?.observe(this.observed);
			this.observer?.observe(view.dom);
		}
		this.observed?.classList.add('dve-equation-editing');
		this.observed?.style.setProperty(
			'--dve-equation-editor-height',
			`${this.element.offsetHeight}px`,
		);
		const width = Math.min(480, Math.max(240, rect.width + 40));
		const left = Math.max(4, Math.min(rect.left - 8, window.innerWidth - width - 8));
		Object.assign(this.element.style, {
			left: `${left}px`,
			top: `${Math.min(rect.bottom + 4, window.innerHeight - 48)}px`,
			width: `${width}px`,
		});
	}

	/** Builds the box's text up (or keeps it linear) and writes it to the equation. */
	commit(): void {
		const view = this.host.view();
		const node = this.pos === undefined ? null : view?.state.doc.nodeAt(this.pos);
		if (!view || !node || this.pos === undefined || !this.edit || this.input.readOnly) return;
		const format = this.host.format();
		const omml = kit().writeEditedEquation({
			text: this.input.value,
			format,
			display: Boolean(node.attrs.display),
			linear: this.edit.linear,
			frame: this.edit.frame,
		});
		this.setStatus(
			omml
				? undefined
				: format === 'latex'
					? 'This LaTeX could not be converted to an equation.'
					: 'This equation is nested too deeply to build up.',
		);
		if (!omml) return;
		this.committed = omml;
		setEquationXml(view, this.pos, omml);
		this.reposition();
	}

	/** Replaces the box's selection with `text` and selects its first placeholder. */
	insert(text: string): void {
		if (!this.editable) return;
		const { input } = this;
		const start = input.selectionStart ?? input.value.length;
		const end = input.selectionEnd ?? start;
		const selected = input.value.slice(start, end);
		// Word puts the selection into the first placeholder of a structure inserted around it.
		const filled =
			selected && text.includes(EQUATION_PLACEHOLDER)
				? text.replace(EQUATION_PLACEHOLDER, () =>
						placeholderArgument(text, selected, this.host.format()),
					)
				: text;
		input.setRangeText(filled, start, end, 'end');
		const placeholder = input.value.indexOf(EQUATION_PLACEHOLDER, start);
		if (placeholder >= 0 && placeholder < start + filled.length)
			input.setSelectionRange(placeholder, placeholder + 1);
		input.focus({ preventScroll: true });
		this.commit();
	}

	/** Normal Text: wraps the selection in normal text, or unwraps the normal text at the caret. */
	toggleNormalText(): void {
		if (!this.editable) return;
		const { input } = this;
		const format = this.host.format();
		const caret = input.selectionStart ?? input.value.length;
		const around = normalTextAround(input.value, caret, format);
		if (around) {
			const [from, to] = around;
			const inner = input.value.slice(from + (format === 'unicode' ? 1 : 6), to - 1);
			input.setRangeText(inner, from, to, 'end');
		} else {
			const end = input.selectionEnd ?? caret;
			const selected = input.value.slice(caret, end);
			const [open, close] = format === 'unicode' ? ['"', '"'] : ['\\text{', '}'];
			input.setRangeText(`${open}${selected}${close}`, caret, end, 'end');
			const inside = caret + open.length + selected.length;
			input.setSelectionRange(inside, inside);
		}
		input.focus({ preventScroll: true });
		this.commit();
	}

	/** Switches the input format, rewriting the box's text in the new notation. */
	setFormat(format: EquationInputFormat): void {
		const old = this.host.format();
		if (old === format) return;
		this.host.setFormat(format);
		const view = this.host.view();
		const node = this.pos === undefined ? null : view?.state.doc.nodeAt(this.pos);
		if (!node) return;
		this.load(String(node.attrs.omml));
		this.input.focus({ preventScroll: true });
	}

	/** Professional or Linear display for this equation. */
	setLinear(linear: boolean): void {
		if (!this.edit || this.pos === undefined || linear === this.edit.linear) return;
		this.edit = { ...this.edit, linear };
		this.commit();
		this.input.focus({ preventScroll: true });
	}

	private exit(after: boolean): void {
		const view = this.host.view();
		const node = this.pos === undefined ? null : view?.state.doc.nodeAt(this.pos);
		if (!view || !node || this.pos === undefined) return;
		const target = after ? this.pos + node.nodeSize : this.pos;
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, target)));
		focusView(view);
	}

	private keydown(event: KeyboardEvent): void {
		// Keys that finish an IME composition belong to the composition.
		if (event.isComposing || event.keyCode === 229) return;
		const { input } = this;
		const start = input.selectionStart ?? 0;
		const end = input.selectionEnd ?? start;
		const mod = event.ctrlKey || event.metaKey;
		if (mod && !event.altKey && /^[zy]$/i.test(event.key)) {
			event.preventDefault();
			const toolbar = this.host.toolbar();
			const redo = event.key.toLowerCase() === 'y' || event.shiftKey;
			if (toolbar) emit(toolbar, 'ribbon-action', { type: 'history', key: redo ? 'redo' : 'undo' });
			return;
		}
		if (event.key === ' ' && this.host.format() === 'unicode' && this.editable) {
			const corrected = kit().mathAutoCorrect(input.value, start);
			if (corrected && start === end) {
				event.preventDefault();
				input.value = corrected.text;
				input.setSelectionRange(corrected.caret, corrected.caret);
				this.commit();
			}
		} else if (event.key === 'Tab' && !mod && !event.altKey) {
			// Tab moves between placeholders and never leaves the equation for the page.
			event.preventDefault();
			const text = input.value;
			const next = event.shiftKey
				? text.lastIndexOf(EQUATION_PLACEHOLDER, Math.max(0, start - 1))
				: text.indexOf(EQUATION_PLACEHOLDER, end);
			const wrapped =
				next >= 0
					? next
					: event.shiftKey
						? text.lastIndexOf(EQUATION_PLACEHOLDER)
						: text.indexOf(EQUATION_PLACEHOLDER);
			if (wrapped >= 0) input.setSelectionRange(wrapped, wrapped + 1);
		} else if (event.key === 'ArrowRight' && start === end && end === input.value.length) {
			event.preventDefault();
			this.exit(true);
		} else if (event.key === 'ArrowLeft' && start === end && start === 0) {
			event.preventDefault();
			this.exit(false);
		} else if (event.key === 'Escape' || event.key === 'Enter') {
			event.preventDefault();
			event.stopPropagation();
			this.exit(true);
		} else if (event.key === 'Backspace' && input.value === '' && this.editable) {
			event.preventDefault();
			const view = this.host.view();
			const pos = this.pos;
			const node = pos === undefined ? null : view?.state.doc.nodeAt(pos);
			if (view?.editable && node && pos !== undefined) {
				view.dispatch(view.state.tr.delete(pos, pos + node.nodeSize));
				focusView(view);
			}
		}
	}
}

export * from './equation-insertions';
