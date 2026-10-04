import { OfficeElement, controlStyles } from './base.js';
import { glyph } from './glyph.js';
import { definer } from './registry.js';
import css from './icon.css?raw';

/** `<office-ui-icon name="check" label="Done">`: decorative unless `label` is set. */
export class OfficeUiIcon extends OfficeElement {
	static override styles = controlStyles(css);
	static override properties = {
		name: { type: String },
		label: { type: String },
	};
	declare name: string;
	declare label: string;

	constructor() {
		super();
		this.name = '';
		this.label = '';
	}

	protected override willUpdate(): void {
		if (this.label) {
			this.setAttribute('role', 'img');
			this.setAttribute('aria-label', this.label);
			this.removeAttribute('aria-hidden');
		} else {
			this.removeAttribute('role');
			this.removeAttribute('aria-label');
			this.setAttribute('aria-hidden', 'true');
		}
	}

	protected override render() {
		return glyph(this.name, 'icon');
	}
}

export const defineIcon = definer('office-ui-icon', () => OfficeUiIcon);
