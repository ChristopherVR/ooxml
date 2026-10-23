import { html, nothing, type TemplateResult } from 'lit';

/** A node's own look, as the per-line style tools show and edit it. */
export interface OfficeSmartArtNodeStyle {
	fillColor?: string;
	fontColor?: string;
	bold?: boolean;
	italic?: boolean;
}

/** `office-smartart-node-move` detail: the node and the direction among its siblings. */
export interface OfficeSmartArtNodeMove {
	id: string;
	direction: -1 | 1;
}

/** `office-smartart-node-style` detail: the node and the part of its style that changed. */
export interface OfficeSmartArtNodeStyleChange {
	id: string;
	style: OfficeSmartArtNodeStyle;
}

/** The strings of the line tools (shown only when the pane's `tools` is set). */
export interface OfficeSmartArtTextPaneToolLabels {
	/** The header button that appends a top-level node. */
	add: string;
	moveUp: string;
	moveDown: string;
	addChild: string;
	remove: string;
	/** Tooltip of the remove button when the node cannot be removed. */
	removeUnavailable: string;
	fill: string;
	fontColor: string;
	bold: string;
	italic: string;
}

/** What the tools need of a line; the flags default to allowed. */
export interface OfficeSmartArtToolNode {
	id: string;
	text: string;
	label?: string;
	style?: OfficeSmartArtNodeStyle;
	canMoveUp?: boolean;
	canMoveDown?: boolean;
	canRemove?: boolean;
	canAddChild?: boolean;
}

/** Swatch values when a node has no colour of its own (Office's accent 1 and white text). */
const DEFAULT_FILL = '#4472c4';
const DEFAULT_FONT = '#ffffff';

type Fire = (name: string, detail: unknown) => void;

/** The buttons after a line's text: move up, move down, add a sub-item and remove. */
export function renderLineActions(
	node: OfficeSmartArtToolNode,
	name: string,
	text: OfficeSmartArtTextPaneToolLabels,
	disabled: boolean,
	fire: Fire,
): TemplateResult {
	const canRemove = node.canRemove !== false;
	const button = (
		className: string,
		label: string,
		glyph: string,
		blocked: boolean,
		onClick: () => void,
		title = label,
	) =>
		html`<button
			type="button"
			class=${`tool ${className}`}
			aria-label=${`${label}: ${name}`}
			title=${title}
			?disabled=${disabled || blocked}
			@click=${onClick}
			>${glyph}</button
		>`;
	return html`<span class="tools"
		>${button('move-up', text.moveUp, '↑', node.canMoveUp === false, () =>
			fire('office-smartart-node-move', { id: node.id, direction: -1 }),
		)}${button('move-down', text.moveDown, '↓', node.canMoveDown === false, () =>
			fire('office-smartart-node-move', { id: node.id, direction: 1 }),
		)}${
			node.canAddChild === false
				? nothing
				: button('add-child', text.addChild, '+', false, () =>
						fire('office-smartart-node-add-child', { id: node.id }),
					)
		}${button(
			'remove',
			text.remove,
			'×',
			!canRemove,
			() => fire('office-smartart-node-remove', { id: node.id }),
			canRemove ? text.remove : text.removeUnavailable,
		)}</span
	>`;
}

/** The style row under a line: fill and text colour swatches, bold and italic toggles. */
export function renderLineStyle(
	node: OfficeSmartArtToolNode,
	name: string,
	text: OfficeSmartArtTextPaneToolLabels,
	disabled: boolean,
	fire: Fire,
): TemplateResult {
	const style = node.style ?? {};
	const change = (patch: OfficeSmartArtNodeStyle) =>
		fire('office-smartart-node-style', { id: node.id, style: patch });
	const swatch = (
		className: string,
		label: string,
		value: string,
		key: 'fillColor' | 'fontColor',
	) =>
		html`<label class=${`swatch ${className}`}
			><span>${label}</span
			><input
				type="color"
				aria-label=${`${label}: ${name}`}
				.value=${value}
				?disabled=${disabled}
				@change=${(event: Event) => change({ [key]: (event.target as HTMLInputElement).value })}
		/></label>`;
	const toggle = (className: string, label: string, glyph: string, key: 'bold' | 'italic') =>
		html`<button
			type="button"
			class=${`tool toggle ${className}`}
			aria-label=${`${label}: ${name}`}
			aria-pressed=${style[key] ? 'true' : 'false'}
			title=${label}
			?disabled=${disabled}
			@click=${() => change({ [key]: !style[key] })}
			>${glyph}</button
		>`;
	return html`<span class="styles" role="group" aria-label=${name}
		>${swatch('fill', text.fill, style.fillColor ?? DEFAULT_FILL, 'fillColor')}${swatch(
			'font',
			text.fontColor,
			style.fontColor ?? DEFAULT_FONT,
			'fontColor',
		)}${toggle('bold', text.bold, 'B', 'bold')}${toggle('italic', text.italic, 'I', 'italic')}</span
	>`;
}
