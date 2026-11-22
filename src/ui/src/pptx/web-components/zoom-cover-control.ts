import type { PptxElement, ZoomPptxElement } from 'ooxml-core/pptx';
import { defineElement } from '../../registry';
import { definePptxSelect } from './office-aliases';
import { attachScopedRibbonStyles } from './ribbon-scoped-styles';
import { ZOOM_COVER_TYPES, zoomCover } from './zoom-cover-image';
import css from './zoom-cover-control.css?raw';

export interface ZoomCoverControlState {
	element: PptxElement | undefined;
	translate: (key: string) => string;
	disabled?: boolean | undefined;
}

export type ZoomCoverPatch = Partial<
	Pick<ZoomPptxElement, 'zoomImageType' | 'imageData' | 'imagePath' | 'summaryTargets'>
>;
export type ZoomCoverPatchEvent = CustomEvent<ZoomCoverPatch>;

export interface PptxUiZoomCoverElement extends HTMLElement {
	state: ZoomCoverControlState;
}

declare global {
	interface HTMLElementTagNameMap {
		'pptx-ui-zoom-cover': PptxUiZoomCoverElement;
	}
}

export function definePptxZoomCover(registry: CustomElementRegistry): void {
	if (registry.get('pptx-ui-zoom-cover')) return;
	definePptxSelect(registry);
	class ZoomCoverControl extends HTMLElement implements PptxUiZoomCoverElement {
		private model: ZoomCoverControlState = { element: undefined, translate: (key) => key };
		private readonly tiles = this.ownerDocument.createElement('pptx-ui-select');
		private readonly replace = this.ownerDocument.createElement('button');
		private readonly reset = this.ownerDocument.createElement('button');
		private readonly file = this.ownerDocument.createElement('input');
		private readonly status = this.ownerDocument.createElement('span');
		private busy = false;
		private tileLabels = '';

		constructor() {
			super();
			this.file.type = 'file';
			this.file.accept = ZOOM_COVER_TYPES;
			this.file.hidden = true;
			this.file.addEventListener('change', () => void this.changeCover());
			this.tiles.addEventListener('change', () => this.sync());
			this.replace.type = 'button';
			this.replace.addEventListener('click', () => this.file.click());
			this.reset.type = 'button';
			this.reset.addEventListener('click', () => this.commit(this.tiles.value));
			this.status.setAttribute('role', 'status');
			this.status.hidden = true;
		}

		get state() {
			return this.model;
		}

		set state(value: ZoomCoverControlState) {
			if (this.model.element?.id !== value.element?.id) {
				this.tiles.value = '';
				this.status.hidden = true;
			}
			this.model = value;
			this.sync();
		}

		connectedCallback() {
			attachScopedRibbonStyles(this.ownerDocument, 'pptx-zoom-cover-styles', this.localName, css);
			this.append(this.tiles, this.replace, this.reset, this.file, this.status);
			this.sync();
		}

		private get zoom() {
			const element = this.model.element;
			return element?.type === 'zoom' ? element : undefined;
		}

		private sync() {
			const element = this.zoom;
			this.hidden = !element;
			if (!element) return;
			const t = this.model.translate;
			const summary = element.zoomType === 'summary';
			const values = summary ? (element.summaryTargets ?? []) : [];
			const options = values.map((tile) => ({
				value: tile.sectionId,
				label: tile.title || `#${tile.targetSlideIndex + 1}`,
			}));
			const labelSignature = JSON.stringify(options);
			if (labelSignature !== this.tileLabels) {
				const previous = this.tiles.value;
				this.tiles.replaceChildren(
					...options.map(({ value, label }) => {
						const option = this.ownerDocument.createElement('option');
						option.value = value;
						option.textContent = label;
						return option;
					}),
				);
				this.tiles.value = options.some(({ value }) => value === previous)
					? previous
					: (options[0]?.value ?? '');
				this.tileLabels = labelSignature;
			}
			this.tiles.hidden = !summary;
			this.tiles.setAttribute('aria-label', t('pptx.zoom.summaryZoom'));
			const target = summary ? values.find((tile) => tile.sectionId === this.tiles.value) : element;
			const disabled = this.busy || this.model.disabled === true || !target;
			this.tiles.disabled = this.busy || this.model.disabled === true;
			this.replace.textContent = t('pptx.graphic.zoomChangeCover');
			this.replace.disabled = disabled;
			this.reset.textContent = t('pptx.graphic.zoomResetCover');
			this.reset.disabled = disabled || target?.zoomImageType !== 'cover';
		}

		private async changeCover() {
			const file = this.file.files?.[0];
			const id = this.zoom?.id;
			if (!file || !id || this.replace.disabled) return;
			const sectionId = this.tiles.value;
			this.file.value = '';
			this.busy = true;
			this.status.hidden = true;
			this.sync();
			try {
				const image = await zoomCover(this.ownerDocument, file);
				if (this.isConnected && this.zoom?.id === id) this.commit(sectionId, image);
			} catch {
				if (this.isConnected && this.zoom?.id === id) {
					this.status.textContent = this.model.translate('pptx.graphic.zoomCoverFailed');
					this.status.hidden = false;
				}
			} finally {
				this.busy = false;
				this.sync();
			}
		}

		private commit(sectionId: string, image?: string) {
			const zoom = this.zoom;
			if (!zoom || this.model.disabled) return;
			const cover = {
				zoomImageType: image === undefined ? ('preview' as const) : ('cover' as const),
				imageData: image,
				imagePath: undefined,
			};
			let patch: ZoomCoverPatch = cover;
			if (zoom.zoomType === 'summary') {
				if (!zoom.summaryTargets?.some((tile) => tile.sectionId === sectionId)) return;
				patch = {
					summaryTargets: zoom.summaryTargets.map((tile) =>
						tile.sectionId === sectionId ? { ...tile, ...cover } : tile,
					),
				};
			}
			this.dispatchEvent(new CustomEvent('zoom-cover-patch', { detail: patch, bubbles: true }));
		}
	}
	defineElement(registry, 'pptx-ui-zoom-cover', ZoomCoverControl);
}
