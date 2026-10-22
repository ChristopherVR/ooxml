/**
 * The review preview's read-only switch for a mounted editor. Visio has no read-only property: its
 * `canEdit` is the switch. The value it had before the review (a shared viewer's read-only role, for
 * one) comes back when the review ends, so ending a review never grants edits the host withheld.
 */
const canEditBeforeReview = new WeakMap();

export function setEditorReadOnly(element, kind, on) {
	if (kind !== 'vsdx') {
		element.readOnly = on;
		return;
	}
	if (on) {
		if (!canEditBeforeReview.has(element)) canEditBeforeReview.set(element, element.canEdit);
		element.canEdit = false;
		return;
	}
	element.canEdit = canEditBeforeReview.get(element) ?? true;
	canEditBeforeReview.delete(element);
}
