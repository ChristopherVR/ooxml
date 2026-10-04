import { present } from './registry.js';
import { tok } from './tokens.js';
import { definer } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

/**
 * A ribbon checkbox row (View > Ruler, Gridlines): a labelled `office-ui-checkbox` sharing its
 * keyboard and focus behaviour. Moved from pptx-viewer `pptx-ui-ribbon-toggle`. Controlled:
 * `label`, `checked`, `disabled`, `title` and `command` attributes; user activation emits
 * `office-toggle` `{ command, checked }` (static `requestEvent`, `idAttribute` and `detailKey`
 * let a product subclass keep its published names). The host reflects the new state.
 */
const CSS = `
:host { display: block; }
label { display: flex; align-items: center; gap: ${tok('--office-space-1')}; min-height: ${tok('--office-ribbon-row-height')}; padding: 0 ${tok('--office-space-1')};
	color: ${tok('--office-foreground')}; font: inherit; font-size: ${tok('--office-font-size-sm')}; white-space: nowrap; cursor: pointer; }
:host([disabled]) label { color: ${tok('--office-muted-foreground')}; cursor: not-allowed; }
@media (pointer: coarse), (max-width: 767px) { label { min-height: ${tok('--office-target-size-touch')}; font-size: ${tok('--office-font-size-sm')}; } }
`;

type Configured = { requestEvent: string; idAttribute: string; detailKey: string };

export const defineRibbonToggle = definer('office-ui-ribbon-toggle', () => {
	class OfficeUiRibbonToggle extends HTMLElement {
		static requestEvent = 'office-toggle';
		static idAttribute = 'command';
		static detailKey = 'command';
		/** The checkbox drawn inside; a product subclass may use its own aliased tag. */
		static checkboxTag = 'office-ui-checkbox';
		static get observedAttributes(): string[] {
			return ['label', 'checked', 'disabled', 'title'];
		}
		private readonly checkbox: HTMLElement & { checked: boolean; disabled: boolean };
		private readonly text: Text;
		constructor() {
			super();
			const doc = this.ownerDocument;
			const root = this.attachShadow({ mode: 'open' });
			attachStyles(root, controlCss(CSS));
			const label = doc.createElement('label');
			this.checkbox = doc.createElement(
				(this.constructor as unknown as { checkboxTag: string }).checkboxTag,
			) as typeof this.checkbox;
			this.text = doc.createTextNode('');
			label.append(this.checkbox, this.text);
			label.addEventListener('click', (event) => {
				event.preventDefault();
				if (event.target !== this.checkbox) this.checkbox.click();
			});
			this.checkbox.addEventListener('input', (event) => event.stopPropagation());
			this.checkbox.addEventListener('change', (event) => {
				event.stopPropagation();
				const checked = this.checkbox.checked;
				this.sync();
				const { requestEvent, idAttribute, detailKey } = this.constructor as unknown as Configured;
				const id = this.getAttribute(idAttribute);
				if (!this.hasAttribute('disabled') && id)
					this.dispatchEvent(
						new CustomEvent(requestEvent, {
							detail: { [detailKey]: id, checked },
							bubbles: true,
							composed: true,
						}),
					);
			});
			root.append(label);
		}
		connectedCallback(): void {
			this.sync();
		}
		attributeChangedCallback(): void {
			this.sync();
		}
		get checked(): boolean {
			return this.hasAttribute('checked');
		}
		set checked(value: boolean) {
			this.toggleAttribute('checked', present(value));
		}
		private sync(): void {
			const label = this.getAttribute('label') ?? '';
			this.text.textContent = label;
			this.checkbox.setAttribute('aria-label', label);
			this.checkbox.title = this.getAttribute('title') ?? label;
			this.checkbox.checked = this.hasAttribute('checked');
			this.checkbox.disabled = this.hasAttribute('disabled');
		}
	}
	return OfficeUiRibbonToggle;
});
