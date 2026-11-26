import { html } from 'lit';
import { repeat } from 'lit/directives/repeat.js';

import { controlStyles, OfficeElement } from '../../base';
import { definer } from '../../registry';
import { animationsLabel } from '../render/ribbon-animations-state';
import { animationPaneFor } from './animation-pane-scope';
import css from './animation-effect-rows.css?raw';

/** Individual selectable effects inside a click-group row in the pane. */
class AnimationEffectRows extends OfficeElement {
	static override watched = ['group-order', 'all'];
	static override styles = controlStyles(css);

	private controller: ReturnType<typeof animationPaneFor>;
	private unsubscribe?: () => void;

	override connectedCallback() {
		this.controller = animationPaneFor(this);
		this.unsubscribe = this.controller?.subscribe(() => this.requestUpdate());
		super.connectedCallback();
	}

	override disconnectedCallback() {
		this.unsubscribe?.();
		this.unsubscribe = undefined;
		this.controller = undefined;
		super.disconnectedCallback();
	}

	override render() {
		const snapshot = this.controller?.getSnapshot();
		const order = this.getAttribute('group-order');
		const clickOrder = order === null ? undefined : Number(order);
		const all = this.hasAttribute('all');
		if (!snapshot || (all && (!snapshot.context?.paneOpen || snapshot.context.hasSelection))) {
			return html`<slot></slot>`;
		}
		const rows = all ? snapshot.rows : snapshot.rows.filter((row) => row.clickOrder === clickOrder);
		if (!rows.length) return html`<slot></slot>`;
		const heading =
			all && snapshot.context
				? html`<h3>${animationsLabel(snapshot.context, 'pptx.animation.timeline', 'Timeline')}</h3>`
				: null;
		return html`${heading}${repeat(
			rows,
			(row) => row.nodeId,
			(row) => html`
				<button
					type="button"
					data-native-effect-id=${row.nodeId}
					aria-pressed=${String(snapshot?.nodeId === row.nodeId)}
					@click=${() => this.controller?.select(row.nodeId)}
					>${row.label}</button
				>
			`,
		)}`;
	}
}

export const definePptxAnimationEffectRows = definer(
	'pptx-ui-animation-effect-rows',
	() => AnimationEffectRows,
);
