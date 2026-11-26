import { defineElement } from '../../registry';
import { canCaptureAnimationPaint } from '../render/animation-painter';
import { canRequestAnimations } from '../render';
import type { RibbonAnimationsIntent, RibbonAnimationsViewState } from '../render';
import { createAnimationPainterController } from '../render/controller/animation-painter-controller';
import { editRibbonAnimationTiming } from '../render/ribbon-animation-editing';
import { animationPaneFor } from './animation-pane-scope';
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
		private animationPane: ReturnType<typeof animationPaneFor>;
		private root: HTMLElement | null = null;
		private unsubscribe?: () => void;
		private unsubscribePane?: () => void;
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
					if (editRibbonAnimationTiming(this.effectiveModel.editing, edit) === 'stale') {
						this.animationPane?.select(undefined);
						if (this.model.editing?.nativeEffectId !== undefined) {
							this.state = {
								...this.model,
								editing: { ...this.model.editing, nativeEffectId: undefined },
							};
						}
					}
					this.render();
				}
			},
			armPainter: (mode) => {
				if (this.model.editable && this.root) {
					this.painter.arm(this.effectiveModel.editing, mode);
				}
			},
		});
		get state() {
			return this.model;
		}
		set state(value: RibbonAnimationsViewState) {
			this.model = value;
			this.animationPane?.sync(value);
			if (!value.editable || !value.editing?.slide || !value.editing.commit) {
				this.painter.cancel();
			}
			this.render();
		}
		connectedCallback(): void {
			attachRibbonAnimationsStyles(this.ownerDocument);
			this.root = this.closest(ANIMATION_PAINTER_SCOPE);
			this.animationPane = animationPaneFor(this);
			this.animationPane?.sync(this.model);
			this.unsubscribePane = this.animationPane?.subscribe(() => this.render());
			if (this.root) {
				this.disconnectInput = connectAnimationPainterInput(this.root, {
					isArmed: () => this.model.editable && this.painter.getSnapshot().mode !== 'off',
					apply: (id) => this.painter.apply(this.effectiveModel.editing, id),
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
			this.unsubscribePane?.();
			this.unsubscribePane = undefined;
			this.animationPane = undefined;
			this.disconnectInput?.();
			this.disconnectInput = undefined;
			this.root = null;
		}
		private get effectiveModel(): RibbonAnimationsViewState {
			const nativeEffectId = this.animationPane?.getSnapshot().nodeId;
			return this.model.editing && nativeEffectId !== undefined
				? { ...this.model, editing: { ...this.model.editing, nativeEffectId } }
				: this.model;
		}
		private render(): void {
			const model = this.effectiveModel;
			this.view.groups.forEach((el, index) => {
				if (this.children[index] !== el) {
					this.insertBefore(el, this.children[index] ?? null);
				}
			});
			const active = this.painter.getSnapshot().mode !== 'off';
			this.view.sync(
				model,
				active,
				active ||
					Boolean(
						this.root &&
						model.editable &&
						model.editing?.commit &&
						model.editing.slide &&
						canCaptureAnimationPaint(
							model.editing.slide,
							model.editing.elementId,
							model.editing.nativeEffectId,
						),
					),
			);
		}
	}
	defineElement(registry, 'pptx-ui-ribbon-animations', RibbonAnimations);
}
