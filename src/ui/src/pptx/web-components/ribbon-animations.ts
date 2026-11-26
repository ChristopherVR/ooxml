import { defineElement } from '../../registry';
import { captureAnimationPaint } from '../render/animation-painter';
import { canRequestAnimations } from '../render';
import type { RibbonAnimationsIntent, RibbonAnimationsViewState } from '../render';
import { createAnimationPainterController } from '../render/controller/animation-painter-controller';
import { editRibbonAnimationTiming } from '../render/ribbon-animation-editing';
import { ANIMATION_PAINTER_SCOPE, connectAnimationPainterInput } from './animation-painter-input';
import { attachRibbonAnimationsStyles } from './ribbon-animations-styles';
import { createRibbonAnimationsView } from './ribbon-animations-view';

export type RibbonAnimationsRequestEvent = CustomEvent<RibbonAnimationsIntent>;
export interface PptxUiRibbonAnimationsElement extends HTMLElement {
	state: RibbonAnimationsViewState;
}
declare global {
	interface HTMLElementTagNameMap {
		'pptx-ui-ribbon-animations': PptxUiRibbonAnimationsElement;
	}
}

/** Controlled Animations chrome; slide edits, history and the pane stay with the host. */
export function definePptxRibbonAnimations(registry: CustomElementRegistry): void {
	if (registry.get('pptx-ui-ribbon-animations')) {
		return;
	}
	class RibbonAnimations extends HTMLElement implements PptxUiRibbonAnimationsElement {
		private model: RibbonAnimationsViewState = { editable: false, hasSelection: false };
		private readonly painter = createAnimationPainterController();
		private selectedEffect: number | undefined;
		private root: HTMLElement | null = null;
		private unsubscribe?: () => void;
		private disconnectInput?: () => void;
		private readonly view = createRibbonAnimationsView(this.ownerDocument, {
			request: (intent) => {
				if (canRequestAnimations(this.model, intent)) {
					this.dispatchEvent(
						new CustomEvent('animations-request', {
							detail: intent,
							bubbles: true,
							composed: true,
						}),
					);
				}
			},
			editTiming: (edit) => {
				if (this.model.editable) {
					editRibbonAnimationTiming(this.model.editing, edit);
					this.render();
				}
			},
			selectEffect: (nodeId) => {
				this.selectedEffect = nodeId;
				this.render();
			},
			armPainter: (mode) => {
				if (this.model.editable && this.root) {
					this.painter.arm(this.model.editing, mode);
				}
			},
		});
		get state() {
			return this.model;
		}
		set state(value: RibbonAnimationsViewState) {
			if (
				value.editing?.slide?.id !== this.model.editing?.slide?.id ||
				value.editing?.elementId !== this.model.editing?.elementId
			)
				this.selectedEffect = undefined;
			this.model = value;
			if (!value.editable || !value.editing?.slide || !value.editing.commit) {
				this.painter.cancel();
			}
			this.render();
		}
		connectedCallback(): void {
			attachRibbonAnimationsStyles(this.ownerDocument);
			this.root = this.closest(ANIMATION_PAINTER_SCOPE);
			if (this.root) {
				this.disconnectInput = connectAnimationPainterInput(this.root, {
					isArmed: () => this.model.editable && this.painter.getSnapshot().mode !== 'off',
					apply: (id) => this.painter.apply(this.model.editing, id),
					cancel: this.painter.cancel,
				});
			}
			this.unsubscribe = this.painter.subscribe(() => {
				this.root?.toggleAttribute(
					'data-pptx-animation-painter',
					this.painter.getSnapshot().mode !== 'off',
				);
				this.render();
			});
			this.render();
		}
		disconnectedCallback(): void {
			this.painter.cancel();
			this.unsubscribe?.();
			this.disconnectInput?.();
			this.disconnectInput = undefined;
			this.view.mountEffectPicker();
			this.root = null;
		}
		private render(): void {
			if (this.model.editing) {
				this.model = {
					...this.model,
					editing: { ...this.model.editing, nativeEffectId: this.selectedEffect },
				};
			}
			this.view.groups.forEach((el, index) => {
				if (this.children[index] !== el) {
					this.insertBefore(el, this.children[index] ?? null);
				}
			});
			const active = this.painter.getSnapshot().mode !== 'off';
			this.view.sync(
				this.model,
				active,
				active ||
					Boolean(
						this.root &&
						this.model.editable &&
						this.model.editing?.commit &&
						this.model.editing.slide &&
						captureAnimationPaint(this.model.editing.slide, this.model.editing.elementId),
					),
			);
			this.mountEffectPicker();
			if (this.model.paneOpen && this.root) queueMicrotask(() => this.mountEffectPicker());
		}
		private mountEffectPicker(): void {
			this.view.mountEffectPicker(
				this.model.paneOpen ? this.root?.querySelector('[data-pptx-effect-picker]') : undefined,
			);
		}
	}
	defineElement(registry, 'pptx-ui-ribbon-animations', RibbonAnimations);
}
