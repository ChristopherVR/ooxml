/** The strings of `<office-ui-cloud-files>` in the editor's locale (File > Open > a cloud place). */
import type { OfficeCloudFilesLabels } from '../../controls';
import type { PageContext } from './parts';

export function cloudLabels(t: PageContext['t']): OfficeCloudFilesLabels {
	return {
		signInHeading: t('Sign in to {provider}'),
		signInText: t('Sign in to open and save your files in {provider}.'),
		signIn: t('Sign in'),
		signOut: t('Sign out'),
		providers: t('Places'),
		recent: t('Recent'),
		myFiles: t('My files'),
		location: t('Location'),
		searchPlaceholder: t('Search files'),
		openLink: t('Open from a link'),
		openLinkPlaceholder: t('Paste a sharing link'),
		openLinkButton: t('Open'),
		files: t('Files'),
		name: t('Name'),
		modified: t('Modified'),
		size: t('Size'),
		loadMore: t('Load more'),
		loading: t('Loading'),
		empty: t('There are no files here.'),
		emptySearch: t('No files match your search.'),
		unsupported: t('This file cannot be opened here'),
		opening: t('Opening {name}'),
		openingPercent: t('Opening {name} ({percent}%)'),
		cancel: t('Cancel'),
		folder: t('Folder'),
		failed: t('Something went wrong.'),
	};
}
