import { html } from 'lit';
import { controlStyles } from '../base';
import { definer } from '../registry';
import { OfficeUiCheckable } from './checkable';
import css from './switch.css?raw';

/**
 * `<office-ui-switch>`: same contract as the checkbox with `role="switch"`. It draws a real knob
 * element, so hosts and tests can measure and style it.
 */
export class OfficeUiSwitch extends OfficeUiCheckable {
	static override styles = controlStyles(css);
	protected readonly semantics = 'switch';

	protected override render() {
		return html`${this.renderCheck()}<span class="knob" part="knob"></span>`;
	}
}

export const defineSwitch = definer('office-ui-switch', () => OfficeUiSwitch);
