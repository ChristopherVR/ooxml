export * from './data-source';
export * from './authoring';
export * from './merge';
export type { MergeFieldInfo } from './list-fields';
export type {
	HeaderFooterSlot,
	HtmlMergeIssue,
	MergeFieldResult,
	MergeFieldsData,
	MergeFieldsOptions,
	MergeFieldsResult,
	MergeRegionInfo,
	MergeRegionResult,
	MergeStory,
} from './headless';
export { mailMergeWithinWorkLimit } from './work-limit';
export type {
	MailMergeIssue,
	MailMergeIssueCode,
	MailMergeOptions,
	MailMergePrompt,
} from './field-values';
export type { AddressBlockOptions, GreetingLineOptions } from './native-fields';
export type {
	MailMergeSettings,
	MailMergeSourceReference,
	MailMergeFieldMapping,
} from './settings';
export type { EnvelopeOptions, LabelOptions, LabelSpec } from './stationery';

export function loadMailMergeStationery() {
	return import('./stationery');
}

export async function resolveMailMergePrompts(
	...args: Parameters<typeof import('./prompts').resolveMailMergePrompts>
) {
	const { resolveMailMergePrompts: resolve } = await import('./prompts');
	return resolve(...args);
}

/** Lists the MERGEFIELD fields and repeating regions of a document; loaded on first use. */
export async function listMergeFields(
	...args: Parameters<typeof import('./list-fields').listMergeFields>
): Promise<ReturnType<typeof import('./list-fields').listMergeFields>> {
	const { listMergeFields: list } = await import('./list-fields');
	return list(...args);
}

/**
 * Fills the MERGEFIELD fields of a document from values, repeating regions and HTML values (see
 * `docs/docx-mail-merge.md`). Loaded on first use: the HTML reader and the region code stay out
 * of this entry point.
 */
export async function mergeFields(
	...args: Parameters<typeof import('./headless').mergeFields>
): Promise<import('./headless').MergeFieldsResult> {
	const { mergeFields: merge } = await import('./headless');
	return merge(...args);
}
