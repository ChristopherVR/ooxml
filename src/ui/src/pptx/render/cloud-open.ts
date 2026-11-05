/**
 * File > Open > cloud places (OneDrive, or a host's own provider), shared by all five bindings.
 *
 * A host passes `fileProviders`; the Open page then lists "This device" plus one place per
 * provider, each provider place rendered by `<pptx-ui-cloud-files>`. A file opened there is loaded
 * through the viewer's normal open path, and this controller remembers it so File > Save, Save As
 * and recovery of THIS document go to `file.storage` (and a file the service leaves read-only
 * stays read-only) until another document is loaded. The bindings only render the places and map
 * {@link resolveCloudDocumentView} onto their own storage, edit and name props.
 *
 * @module render/cloud-open
 */
import type { CloudFileProvider, CloudOpenedFile, DocumentStorage } from 'ooxml-core/embedding';

import { DEFAULT_CLOUD_FILES_LABELS } from '../../files/cloud-files-labels';
import type { OfficeCloudFilesLabels } from '../../files/cloud-files-labels';
import { createControllerStore } from './controller/controller';
import type { PptxController } from './controller/controller';

/** The kinds a PowerPoint File > Open cloud place offers. */
export const CLOUD_OPEN_ACCEPT = ['pptx'] as const;

/** The id of the always-present "This device" place. */
export const DEVICE_OPEN_PLACE_ID = 'device';

/** Matches every binding's translate function (see `BackstageTranslate`). */
type Translate = (key: string, params?: Record<string, string | number>) => string;

export interface OpenPlace {
	/** `device`, or `cloud:<provider id>`. */
	readonly id: string;
	readonly label: string;
	/** Present for a cloud place. */
	readonly provider?: CloudFileProvider;
}

/**
 * The places of File > Open: empty when the host passed no providers (the page stays as it was),
 * otherwise "This device" first and one place per provider.
 */
export function resolveOpenPlaces(
	providers: readonly CloudFileProvider[] | undefined,
	translate: Translate,
): readonly OpenPlace[] {
	if (!providers || providers.length === 0) {
		return [];
	}
	return [
		{ id: DEVICE_OPEN_PLACE_ID, label: translate('pptx.backstage.thisDevice') },
		...providers.map((provider) => ({
			id: `cloud:${provider.id}`,
			label: provider.label,
			provider,
		})),
	];
}

/** The place to show: the selected one, else "This device" (a provider may have been removed). */
export function selectOpenPlace(
	places: readonly OpenPlace[],
	selectedId: string | undefined,
): OpenPlace | undefined {
	return places.find((place) => place.id === selectedId) ?? places[0];
}

/** The element's labels from the viewer dictionary (`pptx.cloudFiles.<label>`), English fallback. */
export function cloudFilesLabels(translate: Translate): OfficeCloudFilesLabels {
	const labels = { ...DEFAULT_CLOUD_FILES_LABELS };
	for (const name of Object.keys(labels) as (keyof OfficeCloudFilesLabels)[]) {
		const key = `pptx.cloudFiles.${name}`;
		// The element fills `{provider}`, `{name}` and `{percent}` itself, so hand them back as is.
		const text = translate(key, {
			provider: '{provider}',
			name: '{name}',
			percent: '{percent}',
		});
		if (text && text !== key) {
			labels[name] = text;
		}
	}
	return labels;
}

export interface CloudDocumentState {
	/** The cloud file now open, or null when the open document did not come from a cloud place. */
	readonly file: CloudOpenedFile | null;
}

export interface CloudDocumentController extends PptxController<CloudDocumentState> {
	/** A cloud file was opened and loaded: its storage and edit right now apply. */
	open(file: CloudOpenedFile): void;
	/** Another document was loaded (new, device, recent or a host content change). */
	clear(): void;
}

export function createCloudDocumentController(): CloudDocumentController {
	const store = createControllerStore<CloudDocumentState>({ file: null });
	return {
		getSnapshot: store.getSnapshot,
		subscribe: store.subscribe,
		open: (file) => store.set({ file }),
		clear: () => {
			if (store.getSnapshot().file) {
				store.set({ file: null });
			}
		},
	};
}

export interface CloudDocumentHostValues {
	storage: DocumentStorage | undefined;
	canEdit: boolean;
	fileName: string | undefined;
	filePath: string | undefined;
}

/**
 * What the viewer uses for storage, edit right, name and recovery key: the host's own, or the
 * open cloud file's. The stricter edit right wins, so a read-only host stays read-only.
 */
export function resolveCloudDocumentView(
	file: CloudOpenedFile | null,
	host: CloudDocumentHostValues,
): CloudDocumentHostValues {
	if (!file) {
		return host;
	}
	return {
		storage: file.storage,
		canEdit: host.canEdit && file.canEdit,
		fileName: file.fileName,
		filePath: `cloud:${file.entry.id}`,
	};
}
