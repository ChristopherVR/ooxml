export const SOURCE_FIELDS = [
	'Title',
	'BookTitle',
	'JournalName',
	'InternetSiteTitle',
	'Publisher',
	'City',
	'Year',
	'Month',
	'Day',
	'Volume',
	'Issue',
	'Pages',
	'Edition',
	'URL',
	'StandardNumber',
] as const;

export const SOURCE_TYPES = [
	'Book',
	'BookSection',
	'JournalArticle',
	'DocumentFromInternetSite',
	'Report',
] as const;
export type SourceField = (typeof SOURCE_FIELDS)[number];

export interface BibliographyPerson {
	first: string;
	middle: string;
	last: string;
}

export type BibliographyAuthor =
	| { kind: 'people'; people: BibliographyPerson[] }
	| { kind: 'corporate'; name: string };

export interface BibliographySource {
	tag: string;
	type: string;
	fields: Partial<Record<SourceField, string>>;
	author?: BibliographyAuthor;
	placeholder?: boolean;
}

export type BibliographyStyle = 'APA' | 'MLA' | 'Chicago';

export interface Bibliography {
	/** The original related custom XML part. */
	part?: string;
	style?: BibliographyStyle;
	sources: BibliographySource[];
}

export const SOURCE_LIMIT = 1000;
export const SOURCE_TEXT_LIMIT = 2048;
