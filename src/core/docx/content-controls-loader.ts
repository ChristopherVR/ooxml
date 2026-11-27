// `fillContentControls`, `readContentControls` and `flattenContentControls` on the
// `ooxml-core/docx` entry. They load the `ooxml-core/docx/content-controls` entry by name on first
// use (it is external to this bundle, as the mail merge is), so this entry, and the ones built on
// it, carry only these loaders. That entry exports the same functions synchronously.
import type { fillContentControls as fill } from './content-controls/fill';
import type { flattenContentControls as flatten } from './content-controls/flatten';
import type { readContentControls as read } from './content-controls/read';
import type { FillContentControlsResult } from './content-controls/types';
import type { DocumentModel } from './model';

/**
 * Fills the tagged content controls of a document: values (text or HTML), repeated and optional
 * controls, locks, and a refill from the template (see `docs/docx-content-controls.md`). Returns a
 * new model for the loaded template's `save()`, with one result per control. Loaded on first use.
 */
export async function fillContentControls(
	...args: Parameters<typeof fill>
): Promise<FillContentControlsResult> {
	const { fillContentControls: run } = await import('ooxml-core/docx/content-controls');
	return run(...args);
}

/** The values of a document's tagged content controls, by tag, as text or HTML. Loaded on first use. */
export async function readContentControls(
	...args: Parameters<typeof read>
): Promise<Record<string, string>> {
	const { readContentControls: run } = await import('ooxml-core/docx/content-controls');
	return run(...args);
}

/** The document with every content control replaced by its content. Loaded on first use. */
export async function flattenContentControls(
	...args: Parameters<typeof flatten>
): Promise<DocumentModel> {
	const { flattenContentControls: run } = await import('ooxml-core/docx/content-controls');
	return run(...args);
}
