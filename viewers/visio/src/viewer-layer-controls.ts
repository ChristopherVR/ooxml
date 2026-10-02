import { getVisioPageLayers } from 'ooxml-core/visio';
import type { ViewerController, ViewerState } from './controller.js';
import { VIEWER_LAYER_LIMITS, layerOverrideMaps } from './viewer-layers.js';

export const layerControlsTemplate = `<details class="layer-controls" hidden><summary>Layers</summary><p>Display changes only. SVG and print snapshots use saved visibility. Guides, NoShow and unknown hidden reasons remain hidden.</p><div data-layer-list></div><p data-layer-status role="status"></p><button type="button" data-layer-reset="all">Reset all layers</button></details>`;

/** Static controls and text-only labels. Document strings never become selectors or markup. */
export function renderLayerControls(panel: HTMLDetailsElement, state: ViewerState): void {
	const model = state.document,
		page = model?.pages[state.pageIndex];
	const list = panel.querySelector<HTMLElement>('[data-layer-list]')!;
	const active = panel.getRootNode() as ShadowRoot;
	const focused = active.activeElement as HTMLElement | null;
	const focus = focused && panel.contains(focused) ? { ...focused.dataset } : undefined;
	const fragment = panel.ownerDocument.createDocumentFragment();
	const overrides = layerOverrideMaps(model, state.layerVisibilityOverrides);
	let total = 0,
		shown = 0;
	for (const source of model && page ? getVisioPageLayers(model, page.id) : []) {
		const ids = new Set<string>();
		const fieldset = panel.ownerDocument.createElement('fieldset');
		const legend = panel.ownerDocument.createElement('legend');
		const pageLabel = `${source.name.slice(0, 256)} (page ${source.id})`;
		legend.textContent = `${pageLabel}${source.id === page!.id ? '' : ' (background)'}`;
		fieldset.append(legend);
		let rows = 0;
		for (const layer of source.layers ?? []) {
			if (ids.has(layer.id)) continue;
			ids.add(layer.id);
			++total;
			if (shown >= VIEWER_LAYER_LIMITS.controls) continue;
			++shown;
			++rows;
			const row = panel.ownerDocument.createElement('div');
			const label = panel.ownerDocument.createElement('label');
			const input = panel.ownerDocument.createElement('input');
			const override = overrides.get(source.id)?.get(layer.id);
			input.type = 'checkbox';
			input.dataset.layerId = layer.id;
			input.dataset.pageId = source.id;
			input.checked = override ?? layer.visible;
			const name = `${layer.name.slice(0, 256) || 'Layer'} (ID ${layer.id})`;
			input.setAttribute('aria-label', `${pageLabel}: ${name}`);
			const description = `${name} · saved ${layer.visible ? 'visible' : 'hidden'}${override === undefined ? '' : ` · override ${override ? 'visible' : 'hidden'}`}`;
			label.append(input, description);
			row.append(label);
			if (override !== undefined) {
				const reset = panel.ownerDocument.createElement('button');
				reset.type = 'button';
				reset.textContent = 'Reset';
				reset.dataset.layerReset = 'one';
				reset.dataset.layerId = layer.id;
				reset.dataset.pageId = source.id;
				reset.setAttribute('aria-label', `Reset ${pageLabel}: ${name}`);
				row.append(reset);
			}
			fieldset.append(row);
		}
		if (rows) fragment.append(fieldset);
	}
	list.replaceChildren(fragment);
	panel.hidden = total === 0 && state.layerVisibilityOverrides.length === 0;
	panel.querySelector('[data-layer-status]')!.textContent =
		total > shown
			? `Showing ${shown} of ${total} layers. Other layers are available through the viewer API.`
			: `${total} layers in this drawing · ${state.layerVisibilityOverrides.length} document display overrides`;
	panel.querySelector<HTMLButtonElement>('[data-layer-reset="all"]')!.disabled =
		state.layerVisibilityOverrides.length === 0;
	if (focus) {
		const controls = [...panel.querySelectorAll<HTMLElement>('input,button')];
		const match = (control: HTMLElement) =>
			control.dataset.layerId === focus.layerId && control.dataset.pageId === focus.pageId;
		const target =
			controls.find(
				(control) => match(control) && control.dataset.layerReset === focus.layerReset,
			) ?? controls.find(match);
		target?.focus();
	}
}

export function wireLayerControls(
	panel: HTMLDetailsElement,
	controller: ViewerController,
): () => void {
	const Abort = panel.ownerDocument.defaultView?.AbortController ?? AbortController;
	const events = new Abort(),
		options = { signal: events.signal };
	panel.addEventListener(
		'change',
		(event) => {
			const input = event.target as HTMLInputElement;
			if (
				input.matches('input[type="checkbox"]') &&
				input.dataset.pageId !== undefined &&
				input.dataset.layerId !== undefined
			)
				controller.setLayerVisibility(input.dataset.pageId, input.dataset.layerId, input.checked);
		},
		options,
	);
	panel.addEventListener(
		'click',
		(event) => {
			const button = (event.target as Element).closest<HTMLButtonElement>(
				'button[data-layer-reset]',
			);
			if (!button || !panel.contains(button)) return;
			if (button.dataset.layerReset === 'all') controller.resetLayerVisibility();
			else if (button.dataset.pageId !== undefined && button.dataset.layerId !== undefined)
				controller.setLayerVisibility(button.dataset.pageId, button.dataset.layerId, null);
		},
		options,
	);
	return () => events.abort();
}
