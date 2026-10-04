import { controlStyles } from '../base.js';
import { definer } from '../registry.js';
import { OfficeUiCheckable } from './checkable.js';
import css from './checkbox.css?raw';

/** `<office-ui-checkbox checked disabled value aria-label>`. */
export class OfficeUiCheckbox extends OfficeUiCheckable {
	static override styles = controlStyles(css);
	protected readonly semantics = 'checkbox';

	protected override render() {
		return this.renderCheck();
	}
}

export const defineCheckbox = definer('office-ui-checkbox', () => OfficeUiCheckbox);
