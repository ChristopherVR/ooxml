// `listMergeFields` and `mergeFields` on the `ooxml-core/docx` entry. They load the
// `ooxml-core/docx/mail-merge` entry by name on first use (it is external to this bundle, as
// `ooxml-core/math` is), so a host that uses both entries shares one copy of the merge and this
// entry, and the ones built on it (`docx/embedded`, `docx/layout`), carry only these loaders.
import type { listMergeFields as list } from './mail-merge/list-fields';
import type { MergeFieldsResult, mergeFields as merge } from './mail-merge/headless';

/** Lists the MERGEFIELD fields and repeating regions of a document; loaded on first use. */
export async function listMergeFields(
	...args: Parameters<typeof list>
): Promise<ReturnType<typeof list>> {
	const { listMergeFields: run } = await import('ooxml-core/docx/mail-merge');
	return run(...args);
}

/**
 * Fills the MERGEFIELD fields of a document from values, repeating regions and HTML values (see
 * `docs/docx-mail-merge.md`). Returns a new model for the loaded template's `save()`, with one
 * result per field. Loaded on first use.
 */
export async function mergeFields(...args: Parameters<typeof merge>): Promise<MergeFieldsResult> {
	const { mergeFields: run } = await import('ooxml-core/docx/mail-merge');
	return run(...args);
}
