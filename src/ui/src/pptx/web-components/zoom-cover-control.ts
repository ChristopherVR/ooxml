import { html, nothing } from 'lit';
import type { PptxElement, ZoomPptxElement } from 'ooxml-core/pptx';
import { controlStyles, OfficeElement } from '../../base';
import { defineElement } from '../../registry';
import { definePptxSelect } from './office-aliases';
import type { PptxUiSelectElement } from './select-value';
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

class ZoomCoverControl extends OfficeElement implements PptxUiZoomCoverElement {
	static override styles = controlStyles(css);
	static override properties = {
		state: { attribute: false, noAccessor: true },
		selectedSection: { state: true },
		busy: { state: true },
		failed: { state: true },
	};

	declare private selectedSection: string;
	declare private busy: boolean;
	declare private failed: boolean;
	private model: ZoomCoverControlState = { element: undefined, translate: (key) => key };

	constructor() {
		super();
		this.selectedSection = '';
		this.busy = false;
		this.failed = false;
	}

	get state(): ZoomCoverControlState {
		return this.model;
	}

	set state(value: ZoomCoverControlState) {
		const previous = this.model;
		this.model = value;
		if (previous.element?.id !== value.element?.id) {
			this.selectedSection = '';
			this.failed = false;
		}
		this.requestUpdate('state', previous);
	}

	private get zoom() {
		const element = this.model.element;
		return element?.type === 'zoom' ? element : undefined;
	}

	private pickCover() {
		this.renderRoot.querySelector<HTMLInputElement>('input')?.click();
	}

	private async changeCover(event: Event, sectionId: string) {
		const input = event.currentTarget as HTMLInputElement;
		const file = input.files?.[0];
		const id = this.zoom?.id;
		input.value = '';
		if (!file || !id || this.busy || this.model.disabled) return;
		this.busy = true;
		this.failed = false;
		try {
			const image = await zoomCover(this.ownerDocument, file);
			if (this.isConnected && this.zoom?.id === id) this.commit(sectionId, image);
		} catch {
			if (this.isConnected && this.zoom?.id === id) this.failed = true;
		} finally {
			this.busy = false;
		}
	}

	private commit(sectionId: string, image?: string) {
		const zoom = this.zoom;
		if (!zoom || this.model.disabled) return;
		const cover: ZoomCoverPatch = {
			zoomImageType: image === undefined ? 'preview' : 'cover',
			imageData: image,
			imagePath: undefined,
		};
		if (zoom.zoomType !== 'summary') {
			this.fire('zoom-cover-patch', cover);
			return;
		}
		if (!zoom.summaryTargets?.some((tile) => tile.sectionId === sectionId)) return;
		this.fire('zoom-cover-patch', {
			summaryTargets: zoom.summaryTargets.map((tile) =>
				tile.sectionId === sectionId ? { ...tile, ...cover } : tile,
			),
		} satisfies ZoomCoverPatch);
	}

	protected override render() {
		const zoom = this.zoom;
		this.toggleAttribute('hidden', !zoom);
		if (!zoom) return nothing;
		const t = this.model.translate;
		const summary = zoom.zoomType === 'summary';
		const tiles = summary ? (zoom.summaryTargets ?? []) : [];
		const tile = tiles.find(({ sectionId }) => sectionId === this.selectedSection) ?? tiles[0];
		const sectionId = tile?.sectionId ?? '';
		const target = summary ? tile : zoom;
		const disabled = this.busy || this.model.disabled === true;
		return html`
			${
				summary
					? html`<pptx-ui-select
							aria-label=${t('pptx.zoom.summaryZoom')}
							.value=${sectionId}
							.disabled=${disabled}
							@change=${(event: Event) => {
								this.selectedSection = (event.currentTarget as PptxUiSelectElement).value;
							}}
							>${tiles.map(
								({ sectionId: id, title, targetSlideIndex }) =>
									html`<option value=${id}>${title || `#${targetSlideIndex + 1}`}</option>`,
							)}</pptx-ui-select
						>`
					: nothing
			}
			<button type="button" ?disabled=${disabled || !target} @click=${this.pickCover}
				>${t('pptx.graphic.zoomChangeCover')}</button
			>
			<button
				type="button"
				?disabled=${disabled || target?.zoomImageType !== 'cover'}
				@click=${() => this.commit(sectionId)}
				>${t('pptx.graphic.zoomResetCover')}</button
			>
			<input
				type="file"
				accept=${ZOOM_COVER_TYPES}
				hidden
				@change=${(event: Event) => void this.changeCover(event, sectionId)}
			/>
			${
				this.failed
					? html`<span role="status">${t('pptx.graphic.zoomCoverFailed')}</span>`
					: nothing
			}
		`;
	}
}

export function definePptxZoomCover(registry: CustomElementRegistry): void {
	if (registry.get('pptx-ui-zoom-cover')) return;
	definePptxSelect(registry);
	defineElement(registry, 'pptx-ui-zoom-cover', ZoomCoverControl);
}
