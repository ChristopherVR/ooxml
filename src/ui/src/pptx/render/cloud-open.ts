/**
 * File > Open > cloud places (OneDrive, or a host's own provider), shared by all five bindings.
 *
 * A host passes `fileProviders`; the Open page then renders `<pptx-ui-open-places>` (the shared
 * `<office-ui-open-places>`), which lists "This device" plus one place per provider. A file opened there is loaded
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
import type { OfficeOpenPlacesLabels } from '../../files/open-places';
import { createControllerStore } from './controller/controller';
import type { PptxController } from './controller/controller';

export type { CloudFileProvider, CloudOpenedFile } from 'ooxml-core/embedding';
export type { OfficeOpenPlacesLabels } from '../../files/open-places';

/** The kinds a PowerPoint File > Open cloud place offers. */
export const CLOUD_OPEN_ACCEPT = ['pptx'] as const;

/** Matches every binding's translate function (see `BackstageTranslate`). */
type Translate = (key: string, params?: Record<string, string | number>) => string;

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

/** The labels of `<pptx-ui-open-places>`: the place tabs plus the cloud browser's own. */
export function openPlacesLabels(translate: Translate): OfficeOpenPlacesLabels {
	return {
		...cloudFilesLabels(translate),
		thisDevice: translate('pptx.backstage.thisDevice'),
		places: translate('pptx.backstage.places'),
	};
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
