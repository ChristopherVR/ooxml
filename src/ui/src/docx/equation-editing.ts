// Editing an equation in place: while an equation is selected, a box under it holds the caret and
// the equation's linear text (Unicode or LaTeX). Typing builds the text up into Office Math as it
// changes, a space after a control word runs math AutoCorrect (`\alpha`), and the Equation tab's
// structures and symbols insert at the caret, selecting the first placeholder. The document shows
// the built-up equation (Professional) or its linear text (Linear).
import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { EQUATION_PLACEHOLDER, type EquationInputFormat } from 'ooxml-core/math';
import { equationKit as kit } from './equation-kit';
import { setEquationXml } from './equation-commands';
import { emit } from './events';
import editorCss from './styles/equation-editor.css?raw';
import { focusView } from './focus-view';
import { translate, type EditorLocale } from './localization';

export interface EquationEditingHost {
	view(): EditorView | undefined;
	toolbar(): HTMLElement | undefined;
	locale(): EditorLocale;
	format(): EquationInputFormat;
	setFormat(format: EquationInputFormat): void;
	refresh(): void;
}

/** Where `text` stands at `caret` inside normal text (`"..."` or `\text{...}`), if it does. */
function normalTextAround(
	text: string,
	caret: number,
	format: EquationInputFormat,
): [number, number] | undefined {
	if (format === 'unicode') {
		const before = text.slice(0, caret);
		const quotes = (before.match(/"/g) ?? []).length;
		if (quotes % 2 === 0) return undefined;
		const start = before.lastIndexOf('"');
		const end = text.indexOf('"', caret);
		return end < 0 ? undefined : [start, end + 1];
	}
	const start = text.lastIndexOf('\\text{', caret);
	if (start < 0) return undefined;
	const end = text.indexOf('}', start);
	return end >= caret - 1 && end >= start + 6 && caret > start ? [start, end + 1] : undefined;
}

/** The galleries' and editor's styles, which load with them rather than with the editor. */
export function equationStyle(): HTMLStyleElement {
	const style = document.createElement('style');
	style.textContent = editorCss;
	return style;
}

export class EquationEditingBox {
	readonly element: HTMLDivElement;
	readonly input: HTMLInputElement;
	private readonly status: HTMLElement;
	private pos: number | undefined;
	/** The markup this box last wrote, so edits from elsewhere (undo, collaboration) reload it. */
	private committed = '';
	/** Whether the equation is shown in Linear rather than built up. */
	private linear = false;
	private frame = 0;
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
	}

	get attachedAt(): number | undefined {
		return this.pos;
	}

	/** Whether the caret is in normal text, for the Normal Text toggle. */
	get normalText(): boolean {
		if (this.pos === undefined) return false;
		const caret = this.input.selectionStart ?? this.input.value.length;
		return Boolean(normalTextAround(this.input.value, caret, this.host.format()));
	}

	/** Opens the box on the equation at `pos`, or brings it up to date with the document. */
	sync(view: EditorView, pos: number | undefined, focus: boolean): void {
		if (pos === undefined) return this.detach();
		const node = view.state.doc.nodeAt(pos);
		if (node?.type.name !== 'equation') return this.detach();
		const omml = String(node.attrs.omml);
		const fresh = pos !== this.pos;
		if (fresh || omml !== this.committed) this.load(omml);
		this.pos = pos;
		this.input.readOnly = !view.editable;
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
		if (this.frame) cancelAnimationFrame(this.frame);
		this.frame = 0;
		window.removeEventListener('scroll', this.reposition, true);
		window.removeEventListener('resize', this.reposition);
		this.element.remove();
	}

	private load(omml: string): void {
		const format = this.host.format();
		const linear = kit().linearDisplayText(omml, format);
		this.linear = linear !== undefined;
		this.input.value = linear ?? kit().linearizeEquation(omml, format);
		this.committed = omml;
		this.status.textContent = '';
	}

	/** The equation's rectangle on screen: in Draft its node, in Print Layout the page selection. */
	private anchorRect(view: EditorView): DOMRect | undefined {
		const dom = this.pos === undefined ? null : view.nodeDOM(this.pos);
		const rect = dom instanceof Element ? dom.getBoundingClientRect() : undefined;
		if (rect && rect.width > 0 && rect.height > 0) return rect;
		const root = view.dom.getRootNode() as ParentNode;
		const boxes = [...root.querySelectorAll<HTMLElement>('.dve-print-selection')]
			.map((box) => box.getBoundingClientRect())
			.filter((box) => box.width > 0);
		return boxes.at(-1);
	}

	private place(): void {
		const view = this.host.view();
		if (!view || view.isDestroyed || this.pos === undefined) return;
		const rect = this.anchorRect(view) ?? view.dom.getBoundingClientRect();
		const width = Math.min(480, Math.max(240, rect.width + 40));
		const left = Math.max(4, Math.min(rect.left, window.innerWidth - width - 8));
		Object.assign(this.element.style, {
			left: `${left}px`,
			top: `${Math.min(rect.bottom + 6, window.innerHeight - 48)}px`,
			width: `${width}px`,
		});
	}

	/** Builds the box's text up (or keeps it linear) and writes it to the equation. */
	commit(): void {
		const view = this.host.view();
		const node = this.pos === undefined ? null : view?.state.doc.nodeAt(this.pos);
		if (!view || !node || this.pos === undefined) return;
		const display = Boolean(node.attrs.display);
		const text = this.input.value;
		const omml = this.linear
			? kit().linearEquationXml(text, display)
			: kit().buildUpEquation(text, this.host.format(), display);
		this.status.textContent = omml
			? ''
			: translate(this.host.locale(), 'This LaTeX could not be converted to an equation.');
		if (!omml) return;
		this.committed = omml;
		setEquationXml(view, this.pos, omml);
		this.reposition();
	}

	/** Replaces the box's selection with `text` and selects its first placeholder. */
	insert(text: string): void {
		const { input } = this;
		const start = input.selectionStart ?? input.value.length;
		const end = input.selectionEnd ?? start;
		const selected = input.value.slice(start, end);
		// Word puts the selection into the first placeholder of a structure inserted around it.
		// In linear text a selection of more than one operand is bracketed to stay one argument.
		const argument =
			this.host.format() === 'unicode' && !/^[\p{L}\p{N}]+$/u.test(selected)
				? `(${selected})`
				: selected;
		const filled =
			selected && text.includes(EQUATION_PLACEHOLDER)
				? text.replace(EQUATION_PLACEHOLDER, argument)
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
		if (this.pos === undefined) return;
		const built = kit().buildUpEquation(this.input.value, old, false);
		if (built) this.input.value = kit().linearizeEquation(built, format);
		if (this.linear) this.commit();
		this.input.focus({ preventScroll: true });
	}

	/** Professional or Linear display for this equation. */
	setLinear(linear: boolean): void {
		if (this.pos === undefined || linear === this.linear) return;
		this.linear = linear;
		this.commit();
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
		if (event.key === ' ' && this.host.format() === 'unicode' && !this.linear) {
			const corrected = kit().mathAutoCorrect(input.value, start);
			if (corrected && start === end) {
				event.preventDefault();
				input.value = corrected.text;
				input.setSelectionRange(corrected.caret, corrected.caret);
				this.commit();
			}
		} else if (event.key === 'Tab' && input.value.includes(EQUATION_PLACEHOLDER)) {
			event.preventDefault();
			const after = input.value.indexOf(EQUATION_PLACEHOLDER, end);
			const next = after >= 0 ? after : input.value.indexOf(EQUATION_PLACEHOLDER);
			input.setSelectionRange(next, next + 1);
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
		} else if (event.key === 'Backspace' && input.value === '') {
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
