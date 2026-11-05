import type { OfficeCloudFilesLabels } from '../files/cloud-files-labels';
import { translate, type EditorLocale, type LocalizationKey } from './localization';

/** The dictionary key of each `<office-ui-cloud-files>` label. */
const LABEL_KEYS = {
	signInHeading: 'cloud.signInHeading',
	signInText: 'cloud.signInText',
	signIn: 'cloud.signIn',
	signOut: 'cloud.signOut',
	providers: 'cloud.providers',
	recent: 'cloud.recent',
	myFiles: 'cloud.myFiles',
	location: 'cloud.location',
	searchPlaceholder: 'cloud.searchPlaceholder',
	openLink: 'cloud.openLink',
	openLinkPlaceholder: 'cloud.openLinkPlaceholder',
	openLinkButton: 'cloud.openLinkButton',
	files: 'cloud.files',
	name: 'cloud.name',
	modified: 'cloud.modified',
	size: 'cloud.size',
	loadMore: 'cloud.loadMore',
	loading: 'cloud.loading',
	empty: 'cloud.empty',
	emptySearch: 'cloud.emptySearch',
	unsupported: 'cloud.unsupported',
	opening: 'cloud.opening',
	openingPercent: 'cloud.openingPercent',
	cancel: 'cloud.cancel',
	folder: 'cloud.folder',
	failed: 'cloud.failed',
} as const satisfies Record<keyof OfficeCloudFilesLabels, LocalizationKey>;

/** Word's display-language labels for the shared cloud files browser. */
export function cloudFileLabels(locale: EditorLocale): OfficeCloudFilesLabels {
	const labels = {} as Record<string, string>;
	for (const [field, key] of Object.entries(LABEL_KEYS)) labels[field] = translate(locale, key);
	return labels as unknown as OfficeCloudFilesLabels;
}
