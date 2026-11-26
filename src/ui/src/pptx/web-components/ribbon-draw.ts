import { defineElement } from '../../registry';
import { canRequestDraw, drawLabel } from '../render';
import type { RibbonDrawIntent, RibbonDrawViewState } from '../render';
import { handleDrawKey } from './ribbon-draw-keyboard';
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
		private readonly loadStatus = this.ownerDocument.createElement('span');
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
			this.loadStatus.className = 'advanced-status';
			this.loadStatus.setAttribute('role', 'status');
			this.loadStatus.hidden = true;
			this.view.trigger.addEventListener('click', (event) => {
				if (!this.model.editable) {
					event.preventDefault();
				}
			});
			this.addEventListener('keydown', (event) => {
				this.escape(event);
				handleDrawKey(this, event);
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
						this.loadStatus.hidden = true;
						this.loadStatus.textContent = '';
						this.advanced = mountRibbonDrawAdvanced(this, () => this.model, this.request);
					}
				})
				.catch(() => {
					if (this.isConnected && !this.hidden) {
						this.loadStatus.hidden = false;
						this.renderLoadStatus();
					}
				})
				.finally(() => {
					this.loading = false;
				});
		}
		private renderLoadStatus(): void {
			if (!this.loadStatus.hidden) {
				this.loadStatus.textContent = drawLabel(
					this.model,
					'pptx.draw.loadFailed',
					'Advanced Draw tools could not load.',
				);
			}
		}
		private render(): void {
			if (!this.contains(this.view.group)) {
				this.append(this.view.group, this.loadStatus);
			}
			this.view.sync(this.model);
			this.renderLoadStatus();
			this.advanced?.sync();
		}
	}
	defineElement(registry, 'pptx-ui-ribbon-draw', RibbonDraw);
}
