import { defineElement } from '../../registry';
import { canRequestDraw } from '../render';
import type { RibbonDrawIntent, RibbonDrawViewState } from '../render';
import { attachRibbonDrawStyles } from './ribbon-draw-styles';
import { createRibbonDrawView } from './ribbon-draw-view';

export type RibbonDrawRequestEvent = CustomEvent<RibbonDrawIntent>;
export interface PptxUiRibbonDrawElement extends HTMLElement {
	state: RibbonDrawViewState;
}
declare global {
	interface HTMLElementTagNameMap {
		'pptx-ui-ribbon-draw': PptxUiRibbonDrawElement;
	}
}

/** Controlled inking chrome, distinct from document gestures and history. */
export function definePptxRibbonDraw(registry: CustomElementRegistry): void {
	if (registry.get('pptx-ui-ribbon-draw')) {
		return;
	}
	class RibbonDraw extends HTMLElement implements PptxUiRibbonDrawElement {
		static observedAttributes = ['hidden'];
		private advanced:
			| ReturnType<(typeof import('./ribbon-draw-advanced'))['mountRibbonDrawAdvanced']>
			| undefined;
		private loading = false;
		private model: RibbonDrawViewState = {
			tool: 'select',
			color: '#000000',
			width: 3,
			editable: false,
		};
		private readonly request = (intent: RibbonDrawIntent) => {
			if (canRequestDraw(this.model, intent)) {
				if (intent.kind === 'tool' || intent.kind === 'preset') this.advanced?.toolChanged();
				this.dispatchEvent(
					new CustomEvent('draw-request', { detail: intent, bubbles: true, composed: true }),
				);
				this.advanced?.sync();
			}
		};
		private readonly view = createRibbonDrawView(this.ownerDocument, this.request);
		private readonly outside = (event: PointerEvent) => {
			if (!this.contains(event.target as Node)) {
				this.view.colors.open = false;
			}
		};
		private readonly position = () => this.view.placePalette();
		private readonly escape = (event: KeyboardEvent) => {
			if (event.key === 'Escape' && this.view.colors.open) {
				event.preventDefault();
				event.stopPropagation();
				this.view.colors.open = false;
				this.view.trigger.focus();
			}
		};
		constructor() {
			super();
			this.view.trigger.addEventListener('click', (event) => {
				if (!this.model.editable) {
					event.preventDefault();
				}
			});
			this.view.group.addEventListener('keydown', (event) => {
				if (event.key === 'Escape') {
					this.escape(event);
				}
				if ([' ', 'Enter', 'ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(event.key)) {
					event.stopPropagation();
				}
			});
			this.addEventListener('keydown', (event) => {
				if (
					[
						' ',
						'Enter',
						'ArrowLeft',
						'ArrowRight',
						'ArrowUp',
						'ArrowDown',
						'Home',
						'End',
						'PageUp',
						'PageDown',
					].includes(event.key)
				)
					event.stopPropagation();
			});
		}
		get state() {
			return this.model;
		}
		set state(value: RibbonDrawViewState) {
			this.model = value;
			this.render();
		}
		connectedCallback(): void {
			attachRibbonDrawStyles(this.ownerDocument);
			this.render();
			this.ownerDocument.addEventListener('pointerdown', this.outside);
			this.ownerDocument.addEventListener('keydown', this.escape);
			this.ownerDocument.defaultView?.addEventListener('resize', this.position);
			this.ownerDocument.addEventListener('scroll', this.position, true);
			this.syncAdvanced();
		}
		disconnectedCallback(): void {
			this.advanced?.dispose();
			this.advanced = undefined;
			this.view.colors.open = false;
			this.ownerDocument.removeEventListener('pointerdown', this.outside);
			this.ownerDocument.removeEventListener('keydown', this.escape);
			this.ownerDocument.defaultView?.removeEventListener('resize', this.position);
			this.ownerDocument.removeEventListener('scroll', this.position, true);
		}
		attributeChangedCallback(): void {
			this.syncAdvanced();
		}
		private syncAdvanced(): void {
			if (this.hidden || !this.isConnected) {
				this.advanced?.dispose();
				this.advanced = undefined;
				return;
			}
			if (this.loading || this.advanced) return;
			this.loading = true;
			void import('./ribbon-draw-advanced')
				.then(({ mountRibbonDrawAdvanced }) => {
					if (this.isConnected && !this.hidden) {
						this.advanced = mountRibbonDrawAdvanced(this, () => this.model, this.request);
					}
				})
				.catch(() => {})
				.finally(() => {
					this.loading = false;
				});
		}
		private render(): void {
			if (!this.contains(this.view.group)) {
				this.append(this.view.group);
			}
			this.view.sync(this.model);
			this.advanced?.sync();
		}
	}
	defineElement(registry, 'pptx-ui-ribbon-draw', RibbonDraw);
}
