export * from './data-source';
export * from './authoring';
export * from './merge';
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
