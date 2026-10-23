import { html, nothing, type PropertyValues } from 'lit';
import { OfficeElement, controlStyles } from '../base';
import { definer } from '../registry';
import css from './smartart-text-pane.css?raw';

/** One line of the outline: a node's text at its depth (0 for a top-level node). */
export interface OfficeSmartArtTextNode {
	id: string;
	text: string;
	level: number;
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

export interface OfficeSmartArtTextPaneLabels {
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
};

/**
 * `<office-ui-smartart-text-pane>`: the SmartArt text pane every product shares. It shows the
 * diagram's nodes as an outline of text lines indented by level and turns keys into requests the
 * host applies to its diagram model (the pane never edits the model itself):
 *
 * - typing fires `office-smartart-text-change` (`{ id, text }`);
 * - Enter fires `office-smartart-node-add` (add a node after this one, same level);
 * - Tab and Shift+Tab fire `office-smartart-node-demote` and `office-smartart-node-promote`;
 * - Backspace in an empty line fires `office-smartart-node-remove`;
 * - focusing a line fires `office-smartart-node-select`; Up and Down move between lines.
 *
 * The node events carry `{ id }`. The host sets `nodes` again after applying a request and
 * `selected` to the node to focus (the new node after Enter, the previous one after Backspace).
 * `closable` adds a close button that fires `office-smartart-text-pane-close`. Strings come
 * translated through `labels`.
 */
export class OfficeUiSmartArtTextPane extends OfficeElement {
	static override styles = controlStyles(css);
	static override properties = {
		closable: { type: Boolean },
		selected: { type: String },
	};
	/** Shows the close button. */
	declare closable: boolean;
	/** The id of the node whose line has (or gets) the focus. */
	declare selected: string | null;

	private list: readonly OfficeSmartArtTextNode[] = [];
	private text: OfficeSmartArtTextPaneLabels = ENGLISH;
	private focusPending = false;

	constructor() {
		super();
		this.closable = false;
		this.selected = null;
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
		const request = (name: string) => {
			event.preventDefault();
			this.focusPending = true;
			this.fire(name, { id: node.id } satisfies OfficeSmartArtNodeRequest);
		};
		if (event.key === 'Enter') request('office-smartart-node-add');
		else if (event.key === 'Tab')
			request(event.shiftKey ? 'office-smartart-node-promote' : 'office-smartart-node-demote');
		else if (event.key === 'Backspace' && input.value === '')
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
		const focusInside = this.renderRoot instanceof ShadowRoot && this.renderRoot.activeElement;
		if ((this.focusPending || focusInside) && (changed.has('selected') || changed.has('nodes'))) {
			this.focusPending = false;
			const target = this.inputs().find((input) => input.dataset.node === this.selected);
			if (
				target &&
				this.renderRoot instanceof ShadowRoot &&
				this.renderRoot.activeElement !== target
			)
				target.focus();
		}
	}

	protected override render() {
		const text = this.text;
		return html`<div class="header"
				><span class="title">${text.title}</span>${
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
							>${this.list.map(
								(node, index) =>
									html`<li
										class="line"
										role="treeitem"
										aria-level=${String(node.level + 1)}
										aria-selected=${node.id === this.selected ? 'true' : 'false'}
										style="--level: ${node.level}"
										><span class="bullet" aria-hidden="true"></span
										><input
											type="text"
											data-node=${node.id}
											placeholder=${text.placeholder}
											aria-label=${node.text || text.placeholder}
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
									/></li>`,
							)}</ul
						>`
			}`;
	}
}

export const defineSmartArtTextPane = definer(
	'office-ui-smartart-text-pane',
	() => OfficeUiSmartArtTextPane,
);
