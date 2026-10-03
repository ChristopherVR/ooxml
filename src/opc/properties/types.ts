/**
 * Package document properties shared by every Office format: core (`docProps/core.xml`,
 * Dublin Core and `cp:`), extended (`docProps/app.xml`) and custom (`docProps/custom.xml`).
 * Absent fields are unset. Dates are kept as the W3CDTF text the part holds (ISO 8601).
 */

/** `docProps/core.xml` (ECMA-376 Part 2, 11). */
export interface CoreProperties {
	title?: string;
	subject?: string;
	/** `dc:creator`: the author. */
	creator?: string;
	/** `cp:keywords`: tags. */
	keywords?: string;
	/** `dc:description`: comments. */
	description?: string;
	lastModifiedBy?: string;
	/** `dcterms:created`, W3CDTF. */
	created?: string;
	/** `dcterms:modified`, W3CDTF. */
	modified?: string;
	category?: string;
	/** `cp:contentStatus` (Draft, Final, ...). */
	contentStatus?: string;
	/** `cp:revision`: the revision number as text. */
	revision?: string;
	/** `dc:language` (RFC 3066 tag). */
	language?: string;
	/** `dc:identifier`. */
	identifier?: string;
	/** `cp:lastPrinted`, W3CDTF. */
	lastPrinted?: string;
	/** `cp:version`. */
	version?: string;
}

/** One `HeadingPairs` entry: a group name (`Worksheets`, `Named Ranges`) and its title count. */
export interface HeadingPair {
	name: string;
	count: number;
}

/** The fields of `docProps/app.xml` this model reads and writes; other elements are preserved. */
export interface AppProperties {
	/** `Application` (`Microsoft Excel`). */
	application?: string;
	/** `AppVersion` (`16.0300`). */
	appVersion?: string;
	company?: string;
	manager?: string;
	/** Base for relative hyperlinks (`HyperlinkBase`). */
	hyperlinkBase?: string;
	template?: string;
	/** `DocSecurity`: 0 none, 1 password protected, 2 read-only recommended, 4 read-only enforced, 8 locked for annotations. */
	docSecurity?: number;
	/** `TotalTime`: editing minutes. */
	totalTime?: number;
	scaleCrop?: boolean;
	linksUpToDate?: boolean;
	sharedDoc?: boolean;
	hyperlinksChanged?: boolean;
	/** `HeadingPairs`: groups of `titlesOfParts`, in order. */
	headingPairs?: HeadingPair[];
	/** `TitlesOfParts`: sheet names, named ranges, slide titles... grouped by `headingPairs`. */
	titlesOfParts?: string[];
}

/** Variant types of a custom property value this model edits. */
export type CustomPropertyType = 'lpwstr' | 'i4' | 'r8' | 'bool' | 'filetime';

interface CustomPropertyBase {
	name: string;
	/** Property id (2 and up); allocated on write when absent or duplicated. */
	pid?: number;
	/** `linkTarget`: the name of a bookmark or defined name the value is linked to. */
	linkTarget?: string;
}

/**
 * A `docProps/custom.xml` property. `raw` keeps a value of any other variant type (`vt:i8`,
 * `vt:lpstr`, `vt:date`, vectors...) as the variant element's XML, written back verbatim.
 */
export type CustomProperty = CustomPropertyBase &
	(
		| { type: 'lpwstr'; value: string }
		| { type: 'i4' | 'r8'; value: number }
		| { type: 'bool'; value: boolean }
		/** W3CDTF / ISO 8601 UTC text (`2026-10-03T00:00:00Z`). */
		| { type: 'filetime'; value: string }
		| { type: 'raw'; xml: string }
	);

/** All three property parts of a package. */
export interface DocumentPropertySet {
	core: CoreProperties;
	app: AppProperties;
	custom: CustomProperty[];
}

/** `fmtid` of user-defined custom properties (FMTID_UserDefinedProperties). */
export const CUSTOM_PROPERTIES_FMTID = '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}';

export const PROPERTY_PART_NAMES = {
	core: 'docProps/core.xml',
	app: 'docProps/app.xml',
	custom: 'docProps/custom.xml',
} as const;

export const PROPERTY_CONTENT_TYPES = {
	core: 'application/vnd.openxmlformats-package.core-properties+xml',
	app: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
	custom: 'application/vnd.openxmlformats-officedocument.custom-properties+xml',
} as const;
