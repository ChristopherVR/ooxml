import { html } from 'lit';
import { OfficeElement, controlStyles } from '../base.js';
import { definer } from '../registry.js';
import css from './ribbon-stack.css?raw';

/**
 * Layout for small ribbon commands: a column (Office's three-row stack of Cut, Copy and Format
 * Painter) or, with `orientation="horizontal"`, a row (Bold, Italic, Underline). Not a group.
 */
export class OfficeUiRibbonStack extends OfficeElement {
	static override styles = controlStyles(css);

	protected override render() {
		return html`<slot></slot>`;
	}
}

export const defineRibbonStack = definer('office-ui-ribbon-stack', () => OfficeUiRibbonStack);
