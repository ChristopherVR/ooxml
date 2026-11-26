import { html, render } from 'lit';
import { OfficeElement, controlStyles } from '../../base';
import { defineElement } from '../../registry';
import { definePptxRibbonCommand } from './office-aliases';
import { definePptxRibbonGallery } from './ribbon-gallery';
import {
	selectedSmartArtCanvasNodes,
	watchSmartArtCanvasSelection,
} from './smartart-canvas-selection';
import type { RibbonGroupId } from '../render/customization/ribbon-control-ids';
import type { PictureRibbonState } from './picture-ribbon-state';
import type { ContextualControls } from './contextual-ribbon-controls';
import css from './contextual-ribbon.css?raw';

type RenderControls = typeof import('./contextual-ribbon-controls').renderContextualControls;

export interface PptxContextualRibbonElement extends HTMLElement {
	group: RibbonGroupId | undefined;
	state: PictureRibbonState | undefined;
}

declare global {
	interface HTMLElementTagNameMap {
		'pptx-ui-contextual-ribbon': PptxContextualRibbonElement;
	}
}

export function definePptxContextualRibbon(registry: CustomElementRegistry): void {
	if (registry.get('pptx-ui-contextual-ribbon')) return;
	definePptxRibbonCommand(registry);
	definePptxRibbonGallery(registry);
	class ContextualRibbon extends OfficeElement implements PptxContextualRibbonElement {
		static override styles = controlStyles(css);
		static override properties = {
			group: { type: String },
			state: { attribute: false, noAccessor: true },
		};
		declare group: RibbonGroupId | undefined;
		private model: PictureRibbonState | undefined;
		private active: AbortController | undefined;
		private renderControls: RenderControls | undefined;
		private canvasSelection: AbortController | undefined;

		get state() {
			return this.model;
		}
		set state(value: PictureRibbonState | undefined) {
			if (value?.element !== this.model?.element || !value?.editable) this.active?.abort();
			this.model = value;
			this.requestUpdate();
		}

		override connectedCallback(): void {
			super.connectedCallback();
			this.canvasSelection = new AbortController();
			watchSmartArtCanvasSelection(this, () => this.requestUpdate(), this.canvasSelection.signal);
			void import('./contextual-ribbon-controls').then(({ renderContextualControls }) => {
				this.renderControls = renderContextualControls;
				if (this.isConnected) this.requestUpdate();
			});
		}

		override disconnectedCallback(): void {
			this.canvasSelection?.abort();
			super.disconnectedCallback();
			this.active?.abort();
		}

		private begin(): AbortSignal {
			this.active?.abort();
			const controller = new AbortController();
			this.active = controller;
			return controller.signal;
		}

		protected override render() {
			return html`<slot></slot>`;
		}

		protected override updated(): void {
			if (!this.group || !this.renderControls || !this.model || !this.isConnected) return;
			const context: ContextualControls = {
				host: this,
				group: this.group,
				state: this.model,
				nodeIds:
					this.model.element?.type === 'smartArt'
						? selectedSmartArtCanvasNodes(this, this.model.element.id)
						: [],
				begin: () => this.begin(),
				close: () => this.active?.abort(),
			};
			// The customization stylesheet queries control ids in the light DOM.
			render(this.renderControls(context), this);
		}
	}
	defineElement(registry, 'pptx-ui-contextual-ribbon', ContextualRibbon);
}
