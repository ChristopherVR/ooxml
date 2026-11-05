/**
 * One File > Open place for a cloud provider, mounted the same way by all five bindings: a host
 * element the binding owns (a ref'd `div`) receives a `<pptx-ui-cloud-files>` showing that one
 * provider, with labels from the viewer dictionary, accepting `.pptx`. A file the user opens is
 * handed to `onOpen`; the binding loads it (see `render/cloud-open`).
 */
import type { CloudFileProvider, CloudOpenedFile } from 'ooxml-core/embedding';

import type { OfficeUiCloudFiles } from '../../controls';
import { CLOUD_OPEN_ACCEPT, cloudFilesLabels } from '../render/cloud-open';
import { registerPptxWebControls } from './index';

export interface CloudFilesPlaceOptions {
	provider: CloudFileProvider;
	/** The binding's translate function. */
	translate: (key: string, params?: Record<string, string | number>) => string;
	/** BCP 47 tag for dates; the viewer's language. */
	locale?: string | undefined;
	/** The user opened `file`; load it. */
	onOpen: (file: CloudOpenedFile) => void;
}

export interface CloudFilesPlace {
	/** Re-apply options after the provider, language or callback changed. */
	update(options: CloudFilesPlaceOptions): void;
	destroy(): void;
}

/** Mount a provider's file browser into `host`. */
export function mountCloudFilesPlace(
	host: HTMLElement,
	options: CloudFilesPlaceOptions,
): CloudFilesPlace {
	registerPptxWebControls();
	const element = document.createElement('pptx-ui-cloud-files') as OfficeUiCloudFiles;
	element.setAttribute('data-pptx-cloud-files', '');
	let current = options;
	const apply = (next: CloudFilesPlaceOptions): void => {
		current = next;
		element.providers = [next.provider];
		element.accept = CLOUD_OPEN_ACCEPT;
		element.labels = cloudFilesLabels(next.translate);
		element.locale = next.locale;
	};
	const onOpen = (event: Event): void => {
		const detail = (event as CustomEvent<{ file: CloudOpenedFile }>).detail;
		current.onOpen(detail.file);
	};
	element.addEventListener('office-cloud-file-open', onOpen);
	apply(options);
	host.append(element);
	return {
		update: apply,
		destroy() {
			element.removeEventListener('office-cloud-file-open', onOpen);
			element.remove();
		},
	};
}
