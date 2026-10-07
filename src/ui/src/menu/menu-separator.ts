import { OfficeElement, controlStyles } from '../base';
import { definer } from '../registry';
import css from './menu-separator.css?raw';

/** A `role="separator"` rule between groups of menu items. */
export class OfficeUiMenuSeparator extends OfficeElement {
	static override styles = controlStyles(css);

	protected override willUpdate(): void {
		this.setAttribute('role', 'separator');
	}
}

export const defineMenuSeparator = definer('office-ui-menu-separator', () => OfficeUiMenuSeparator);
