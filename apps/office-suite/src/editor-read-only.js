/**
 * The review preview's read-only switch for a mounted editor. Every full editor has `readOnly`, the
 * host's own switch: a collaboration viewer role is tracked apart from it, so ending a review
 * never grants edits a role withheld and nothing has to be saved and restored.
 */
export function setEditorReadOnly(element, _kind, on) {
	element.readOnly = on;
}
