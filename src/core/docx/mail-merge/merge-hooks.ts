// How the headless merge (`mergeFields`) changes the projection the editor's merge uses.
import type { MailMergeRecord } from './data-source';

/** What happened to one MERGEFIELD, for the headless merge's report. */
export interface MergeFieldEvent {
	blockId: string;
	/** The column the field reads, or the field type when it could not be read. */
	name: string;
	status: 'filled' | 'missing' | 'invalid' | 'locked' | 'unsupported';
}

/** How the headless merge (`mergeFields`) changes the projection the editor's merge uses. */
export interface MergeHooks {
	/** Keep each field and give it a new cached result instead of replacing it by its result. */
	keepFields?: boolean;
	/** The record the fields of a paragraph read; replaces the recipient list. */
	record?(blockId: string): MailMergeRecord;
	/** `keep`: a field the record does not provide keeps the result it has. */
	missing?: 'empty' | 'keep';
	observe?(event: MergeFieldEvent): void;
	/** Whether a result must replace its field even when fields are kept. */
	unlink?(text: string): boolean;
}
