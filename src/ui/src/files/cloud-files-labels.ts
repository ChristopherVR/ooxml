import type { CloudEntry, CloudFileKind } from 'ooxml-core/embedding';

/** Every user-visible string of `<office-ui-cloud-files>`; set `labels` (partial) to localise. */
export interface OfficeCloudFilesLabels {
	/** Heading of the signed-out state; `{provider}` is the provider label. */
	signInHeading: string;
	signInText: string;
	signIn: string;
	signOut: string;
	/** Name of the provider switcher when more than one provider is given. */
	providers: string;
	recent: string;
	myFiles: string;
	/** Name of the breadcrumb trail. */
	location: string;
	searchPlaceholder: string;
	openLink: string;
	openLinkPlaceholder: string;
	openLinkButton: string;
	files: string;
	name: string;
	modified: string;
	size: string;
	loadMore: string;
	loading: string;
	empty: string;
	emptySearch: string;
	/** Tooltip of a file the host does not accept. */
	unsupported: string;
	/** Status while a file downloads; `{name}` is its name. */
	opening: string;
	/** Status with a percentage; `{name}` and `{percent}`. */
	openingPercent: string;
	cancel: string;
	folder: string;
	failed: string;
}

export const DEFAULT_CLOUD_FILES_LABELS: OfficeCloudFilesLabels = {
	signInHeading: 'Sign in to {provider}',
	signInText: 'Sign in to open and save your files in {provider}.',
	signIn: 'Sign in',
	signOut: 'Sign out',
	providers: 'Places',
	recent: 'Recent',
	myFiles: 'My files',
	location: 'Location',
	searchPlaceholder: 'Search files',
	openLink: 'Open from a link',
	openLinkPlaceholder: 'Paste a sharing link',
	openLinkButton: 'Open',
	files: 'Files',
	name: 'Name',
	modified: 'Modified',
	size: 'Size',
	loadMore: 'Load more',
	loading: 'Loading',
	empty: 'There are no files here.',
	emptySearch: 'No files match your search.',
	unsupported: 'This file cannot be opened here',
	opening: 'Opening {name}',
	openingPercent: 'Opening {name} ({percent}%)',
	cancel: 'Cancel',
	folder: 'Folder',
	failed: 'Something went wrong.',
};

export function fillLabel(template: string, values: Record<string, string>): string {
	return template.replace(/\{(\w+)\}/g, (whole, key: string) => values[key] ?? whole);
}

const KIND_BADGE: Record<CloudFileKind, string> = { docx: 'W', xlsx: 'X', pptx: 'P', vsdx: 'V' };

/** The letter in a file's icon: its Office product, nothing for a folder, a dot for anything else. */
export function entryBadge(entry: CloudEntry): string {
	return entry.type === 'folder' ? '' : entry.kind ? KIND_BADGE[entry.kind] : '•';
}

export function formatSize(bytes: number | undefined): string {
	if (bytes === undefined) return '';
	if (bytes < 1024) return `${bytes} B`;
	const units = ['KB', 'MB', 'GB', 'TB'];
	let value = bytes / 1024;
	let unit = 0;
	while (value >= 1024 && unit < units.length - 1) {
		value /= 1024;
		unit++;
	}
	return `${value >= 10 ? Math.round(value) : Math.round(value * 10) / 10} ${units[unit]}`;
}

export function formatModified(modified: number | undefined, locale?: string): string {
	if (modified === undefined) return '';
	try {
		return new Intl.DateTimeFormat(locale, { dateStyle: 'medium' }).format(modified);
	} catch {
		return new Date(modified).toISOString().slice(0, 10);
	}
}

export function messageOf(error: unknown, fallback: string): string {
	return error instanceof Error && error.message ? error.message : fallback;
}
