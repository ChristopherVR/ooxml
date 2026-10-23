import { html, nothing, type PropertyValues } from 'lit';
import { OfficeElement, controlStyles } from '../base';
import { definer } from '../registry';
import css from './smartart-text-pane.css?raw';
import {
	renderLineActions,
	renderLineStyle,
	type OfficeSmartArtNodeStyle,
	type OfficeSmartArtTextPaneToolLabels,
} from './smartart-text-pane-tools';

export type {
	OfficeSmartArtNodeMove,
	OfficeSmartArtNodeStyle,
	OfficeSmartArtNodeStyleChange,
	OfficeSmartArtTextPaneToolLabels,
} from './smartart-text-pane-tools';

/** One line of the outline: a node's text at its depth (0 for a top-level node). */
export interface OfficeSmartArtTextNode {
	id: string;
	text: string;
	level: number;
	/** Shown instead of the round bullet, e.g. the item number ("1") of a top-level line. */
	marker?: string;
	/** Accessible name of the line (default: its text, or the placeholder). */
	label?: string;
	/** The node's own look, shown and edited by the style tools. */
	style?: OfficeSmartArtNodeStyle;
	/** Tool availability; each defaults to allowed. `canRemove` also gates Backspace and Delete. */
	canMoveUp?: boolean;
	canMoveDown?: boolean;
	canRemove?: boolean;
	/** False hides the add sub-item tool for this line. */
	canAddChild?: boolean;
}

/** `office-smartart-text-change` detail. */
export interface OfficeSmartArtTextChange {
	id: string;
	text: string;
}

/**
 * Detail of the node events (`office-smartart-node-add`, `-remove`, `-promote`, `-demote`,
 * `-select`): the node the key was pressed on.
 */
export interface OfficeSmartArtNodeRequest {
	id: string;
}

export interface OfficeSmartArtTextPaneLabels extends OfficeSmartArtTextPaneToolLabels {
	title: string;
	/** Shown in an empty line, as Office shows `[Text]`. */
	placeholder: string;
	close: string;
	/** Accessible name of the outline. */
	outline: string;
	/** Shown when the diagram has no nodes. */
	empty: string;
}

const ENGLISH: OfficeSmartArtTextPaneLabels = {
	title: 'Type your text here',
	placeholder: '[Text]',
	close: 'Close',
	outline: 'SmartArt text',
	empty: 'This diagram has no shapes with text.',
	add: 'Add shape',
	moveUp: 'Move up',
	moveDown: 'Move down',
	addChild: 'Add sub-item',
	remove: 'Remove',
	removeUnavailable: 'The layout needs this shape',
	fill: 'Fill',
	fontColor: 'Text',
	bold: 'Bold',
	italic: 'Italic',
};

/**
 * `<office-ui-smartart-text-pane>`: the SmartArt text pane every product shares. It shows the
 * diagram's nodes as an outline of text lines indented by level and turns keys into requests the
 * host applies to its diagram model (the pane never edits the model itself):
 *
 * - typing fires `office-smartart-text-change` (`{ id, text }`);
 * - Enter fires `office-smartart-node-add` (add a node after this one, same level);
 * - Tab and Shift+Tab fire `office-smartart-node-demote` and `office-smartart-node-promote`;
 * - Backspace or Delete in an empty line fires `office-smartart-node-remove` (unless the line's
 *   `canRemove` is false);
 * - focusing a line fires `office-smartart-node-select`; Up and Down move between lines.
 *
 * The node events carry `{ id }`. The host sets `nodes` again after applying a request and
 * `selected` to the node to focus (the new node after Enter, the previous one after Backspace).
 * `closable` adds a close button that fires `office-smartart-text-pane-close`. Strings come
 * translated through `labels`.
 *
 * `tools` adds PowerPoint's inspector tools: a header button that fires
 * `office-smartart-node-append`, and per line move up and down (`office-smartart-node-move`,
 * `{ id, direction }`), add sub-item (`office-smartart-node-add-child`), remove
 * (`office-smartart-node-remove`) and a style row (fill and text colour, bold, italic) that fires
 * `office-smartart-node-style` (`{ id, style }`). `disabled` makes the pane read-only, `notes`
 * adds hint lines under the outline and `appendDisabled` disables the header button.
 */
export class OfficeUiSmartArtTextPane extends OfficeElement {
	static override styles = controlStyles(css);
	static override properties = {
		closable: { type: Boolean },
		selected: { type: String },
		disabled: { type: Boolean, reflect: true },
		tools: { type: Boolean, reflect: true },
		appendDisabled: { type: Boolean, attribute: 'append-disabled' },
	};
	/** Shows the close button. */
	declare closable: boolean;
	/** Read-only: lines and tools are disabled and keys request nothing. */
	declare disabled: boolean;
	/** Shows the append button and the per-line tools. */
	declare tools: boolean;
	/** Disables the append button (e.g. the layout holds no more shapes). */
	declare appendDisabled: boolean;
	/** The id of the node whose line has (or gets) the focus. */
	declare selected: string | null;

	private list: readonly OfficeSmartArtTextNode[] = [];
	private text: OfficeSmartArtTextPaneLabels = ENGLISH;
	private hints: readonly string[] = [];
	private focusPending = false;
	/** Tracked from focus events: reading a shadow root's `activeElement` is not safe everywhere. */
	private focusWithin = false;

	constructor() {
		super();
		this.closable = false;
		this.selected = null;
		this.disabled = false;
		this.tools = false;
		this.appendDisabled = false;
		this.addEventListener('focusin', () => (this.focusWithin = true));
		this.addEventListener('focusout', () => (this.focusWithin = false));
	}

	/** Hint lines shown under the outline. */
	get notes(): readonly string[] {
		return this.hints;
	}
	set notes(value: readonly string[]) {
		this.hints = value;
		this.requestUpdate('notes');
	}

	get nodes(): readonly OfficeSmartArtTextNode[] {
		return this.list;
	}
	set nodes(value: readonly OfficeSmartArtTextNode[]) {
		this.list = value;
		this.requestUpdate('nodes');
	}
	get labels(): OfficeSmartArtTextPaneLabels {
		return this.text;
	}
	set labels(value: Partial<OfficeSmartArtTextPaneLabels>) {
		this.text = { ...ENGLISH, ...value };
		this.requestUpdate('labels');
	}

	/** Moves the focus to the selected node's line (or the first line). */
	focusSelected(): void {
		const inputs = this.inputs();
		(inputs.find((input) => input.dataset.node === this.selected) ?? inputs[0])?.focus();
	}

	private inputs(): HTMLInputElement[] {
		return [...this.renderRoot.querySelectorAll<HTMLInputElement>('input[data-node]')];
	}

	private onKey(event: KeyboardEvent, node: OfficeSmartArtTextNode, index: number): void {
		const input = event.target as HTMLInputElement;
		if (this.disabled) return;
		const request = (name: string) => {
			event.preventDefault();
			this.focusPending = true;
			this.fire(name, { id: node.id } satisfies OfficeSmartArtNodeRequest);
		};
		if (event.key === 'Enter') request('office-smartart-node-add');
		else if (event.key === 'Tab')
			request(event.shiftKey ? 'office-smartart-node-promote' : 'office-smartart-node-demote');
		else if (
			(event.key === 'Backspace' || event.key === 'Delete') &&
			input.value === '' &&
			node.canRemove !== false
		)
			request('office-smartart-node-remove');
		else if (event.key === 'ArrowUp' || event.key === 'ArrowDown') {
			const next = this.inputs()[index + (event.key === 'ArrowUp' ? -1 : 1)];
			if (!next) return;
			event.preventDefault();
			next.focus();
		}
	}

	protected override updated(changed: PropertyValues<this>): void {
		// Values are written here, not bound, so a line being typed in keeps its caret.
		const byId = new Map(this.list.map((node) => [node.id, node]));
		for (const input of this.inputs()) {
			const text = byId.get(input.dataset.node ?? '')?.text ?? '';
			if (input.value !== text) input.value = text;
		}
		if (
			(this.focusPending || this.focusWithin) &&
			(changed.has('selected') || changed.has('nodes'))
		) {
			this.focusPending = false;
			// Focusing the line that already has the focus is a no-op.
			this.inputs()
				.find((input) => input.dataset.node === this.selected)
				?.focus();
		}
	}

	protected override render() {
		const text = this.text;
		const fire = (name: string, detail: unknown) => this.fire(name, detail);
		return html`<div class="header"
				><span class="title">${text.title}</span>${
					this.tools
						? html`<button
								type="button"
								class="append"
								?disabled=${this.disabled || this.appendDisabled}
								@click=${() => this.fire('office-smartart-node-append', {})}
								>${text.add}</button
							>`
						: nothing
				}${
					this.closable
						? html`<button
								type="button"
								class="close"
								aria-label=${text.close}
								title=${text.close}
								@click=${() => this.fire('office-smartart-text-pane-close', {})}
								>×</button
							>`
						: nothing
				}</div
			>${
				this.list.length === 0
					? html`<p class="empty">${text.empty}</p>`
					: html`<ul class="outline" role="tree" aria-label=${text.outline}
							>${this.list.map((node, index) => {
								const name = node.label ?? (node.text || text.placeholder);
								return html`<li
									class="line"
									role="treeitem"
									aria-level=${String(node.level + 1)}
									aria-selected=${node.id === this.selected ? 'true' : 'false'}
									style="--level: ${node.level}"
									>${
										node.marker === undefined
											? html`<span class="bullet" aria-hidden="true"></span>`
											: html`<span class="marker" aria-hidden="true">${node.marker}</span>`
									}<input
										type="text"
										data-node=${node.id}
										data-testid="smartart-node-text"
										placeholder=${text.placeholder}
										aria-label=${name}
										?disabled=${this.disabled}
										@focus=${() => {
											if (this.selected === node.id) return;
											this.selected = node.id;
											this.fire('office-smartart-node-select', { id: node.id });
										}}
										@input=${(event: Event) =>
											this.fire('office-smartart-text-change', {
												id: node.id,
												text: (event.target as HTMLInputElement).value,
											} satisfies OfficeSmartArtTextChange)}
										@keydown=${(event: KeyboardEvent) => this.onKey(event, node, index)}
									/>${
										this.tools
											? [
													renderLineActions(node, name, text, this.disabled, fire),
													renderLineStyle(node, name, text, this.disabled, fire),
												]
											: nothing
									}</li
								>`;
							})}</ul
						>`
			}${this.hints.map((note) => html`<p class="note" role="note">${note}</p>`)}`;
	}
}

export const defineSmartArtTextPane = definer(
	'office-ui-smartart-text-pane',
	() => OfficeUiSmartArtTextPane,
);
