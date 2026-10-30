// While a restriction is enforced Word greys out the commands it forbids. This does the same to
// the ribbon: editing tabs are disabled (the viewing commands stay), and Track Changes and
// Accept/Reject are disabled when tracked changes protection owns them.
import type { DocumentProtection } from 'ooxml-core/docx';
import { editRestrictionOf } from 'ooxml-core/docx/ui';
import { findLocalizedControl, translate, type EditorLocale } from './localization';

const EDITING_PANELS = ['Home', 'Insert', 'Layout', 'References', 'Mailings', 'Table'];
/** Commands that still work: they read, copy or search, or change nothing in the document. */
const STAYS_ON = new Set(['Copy', 'Select all', 'Find and replace', 'Show paragraph marks']);
const REVIEW_LABELS = ['Track changes', 'Accept', 'Reject', 'Accept all', 'Reject all'];
const REASON = 'Editing is restricted by document protection';

const TITLE = 'data-protection-title';

/** Disables the commands the enforced protection forbids, and enables nothing it did not disable. */
export function syncProtectionControls(
	toolbar: HTMLElement,
	protection: DocumentProtection | undefined,
	locale: EditorLocale,
): void {
	const restriction = editRestrictionOf(protection);
	const lock = (control: HTMLElement) => {
		(control as HTMLButtonElement).disabled = true;
		if (!control.hasAttribute(TITLE)) control.setAttribute(TITLE, control.title);
		control.title = translate(locale, REASON);
	};
	// Put back the tooltips of controls an earlier protection disabled.
	for (const control of toolbar.querySelectorAll<HTMLElement>(`[${TITLE}]`)) {
		control.title = control.getAttribute(TITLE) ?? '';
		control.removeAttribute(TITLE);
	}
	if (!restriction) return;
	if (restriction !== 'trackedChanges')
		for (const panel of EDITING_PANELS)
			for (const control of toolbar.querySelectorAll<HTMLElement>(
				`.ribbon-panel[data-panel="${panel}"] button, .ribbon-panel[data-panel="${panel}"] select, .ribbon-panel[data-panel="${panel}"] input`,
			)) {
				const label = control.dataset.localearialabel ?? control.getAttribute('aria-label') ?? '';
				if (!STAYS_ON.has(label)) lock(control);
			}
	// Tracked changes protection owns Track Changes and the accept/reject commands; a read only
	// document takes no comments either.
	const reviewLabels =
		restriction === 'readOnly'
			? [...REVIEW_LABELS, 'Add comment', 'Delete comment']
			: REVIEW_LABELS;
	for (const label of reviewLabels) {
		const control = findLocalizedControl<HTMLButtonElement>(toolbar, label);
		if (control) lock(control);
	}
}
