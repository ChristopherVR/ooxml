import { html, nothing } from 'lit';
import type { CloudFileKind, CloudFileProvider } from 'ooxml-core/embedding';
import { OfficeElement, controlStyles } from '../base';
import { definer } from '../registry';
import { defineCloudFiles } from './cloud-files';
import type { OfficeCloudFilesLabels } from './cloud-files-labels';
import css from './open-places.css?raw';

/** The labels of the place switcher; the browser's own labels pass through `labels` too. */
export interface OfficeOpenPlacesLabels extends OfficeCloudFilesLabels {
	/** The tab of the product's own device content (Browse, recent files). */
	thisDevice: string;
	/** Accessible name of the place switcher. */
	places: string;
}

export const DEFAULT_OPEN_PLACES_LABELS: Pick<OfficeOpenPlacesLabels, 'thisDevice' | 'places'> = {
	thisDevice: 'This device',
	places: 'Places',
};

/** The id of the device place; every other place id is a provider id. */
export const DEVICE_PLACE = 'device';

export type OfficeOpenPlaceChangeEvent = CustomEvent<{ place: string }>;

const KEY_STEPS: Record<string, number> = {
	ArrowLeft: -1,
	ArrowRight: 1,
	ArrowUp: -1,
	ArrowDown: 1,
};

/**
 * `<office-ui-open-places>`: the places of File > Open, the same in every product. A tab per
 * place: "This device" shows the element's light-DOM children (the product's Browse button and
 * recent list, unchanged), and each `CloudFileProvider` shows `<office-ui-cloud-files>` for it.
 * Without providers it shows only its children, with no tabs. Properties: `providers`, `accept`,
 * `labels` (`thisDevice`, `places` and every cloud browser label), `locale`, `selected` (a place
 * id: `device` or a provider id). Choosing a tab emits `office-open-place-change` `{ place }`;
 * opening a file bubbles the browser's `office-cloud-file-open` `{ providerId, file }` unchanged.
 */
export class OfficeUiOpenPlaces extends OfficeElement {
	static override styles = controlStyles(css);
	static override properties = {
		providers: { attribute: false },
		accept: { attribute: false },
		labels: { attribute: false },
		locale: { type: String },
		selected: { type: String },
	};
	declare providers: readonly CloudFileProvider[];
	declare accept: readonly CloudFileKind[] | undefined;
	declare labels: Partial<OfficeOpenPlacesLabels>;
	declare locale: string | undefined;
	declare selected: string;

	constructor() {
		super();
		this.providers = [];
		this.accept = undefined;
		this.labels = {};
		this.locale = undefined;
		this.selected = DEVICE_PLACE;
	}

	private get places(): { id: string; label: string }[] {
		const thisDevice = this.labels.thisDevice ?? DEFAULT_OPEN_PLACES_LABELS.thisDevice;
		return [
			{ id: DEVICE_PLACE, label: thisDevice },
			...this.providers.map((provider) => ({ id: provider.id, label: provider.label })),
		];
	}

	/** The place shown: the chosen one while it still exists, else This device. */
	get place(): string {
		return this.places.some((place) => place.id === this.selected) ? this.selected : DEVICE_PLACE;
	}

	private choose(id: string, focus = false): void {
		if (id !== this.place) {
			this.selected = id;
			this.fire('office-open-place-change', { place: id });
		}
		if (focus)
			void this.updateComplete.then(() =>
				this.renderRoot.querySelector<HTMLElement>(`[data-place="${CSS.escape(id)}"]`)?.focus(),
			);
	}

	private onKeydown(event: KeyboardEvent): void {
		const { places } = this;
		const index = places.findIndex((place) => place.id === this.place);
		const next =
			event.key === 'Home'
				? 0
				: event.key === 'End'
					? places.length - 1
					: event.key in KEY_STEPS
						? (index + KEY_STEPS[event.key]! + places.length) % places.length
						: -1;
		if (next < 0) return;
		event.preventDefault();
		this.choose(places[next]!.id, true);
	}

	protected override render() {
		if (!this.providers.length) return html`<slot></slot>`;
		const current = this.place;
		const provider = this.providers.find((candidate) => candidate.id === current);
		return html`<div
				class="places"
				role="tablist"
				aria-label=${this.labels.places ?? DEFAULT_OPEN_PLACES_LABELS.places}
				@keydown=${this.onKeydown}
				>${this.places.map(
					(place) =>
						html`<button
							type="button"
							role="tab"
							id=${`place-${place.id}`}
							data-place=${place.id}
							aria-selected=${place.id === current ? 'true' : 'false'}
							aria-controls="panel"
							tabindex=${place.id === current ? 0 : -1}
							@click=${() => this.choose(place.id)}
							>${place.label}</button
						>`,
				)}</div
			><div id="panel" class="panel" role="tabpanel" aria-labelledby=${`place-${current}`}
				>${
					provider
						? html`<office-ui-cloud-files
								.providers=${[provider]}
								.accept=${this.accept ?? ['docx', 'xlsx', 'pptx', 'vsdx']}
								.labels=${this.labels}
								.locale=${this.locale}
							></office-ui-cloud-files>`
						: nothing
				}<slot ?hidden=${provider !== undefined}></slot
			></div>`;
	}
}

export const defineOpenPlaces = definer('office-ui-open-places', () => OfficeUiOpenPlaces, [
	defineCloudFiles,
]);
