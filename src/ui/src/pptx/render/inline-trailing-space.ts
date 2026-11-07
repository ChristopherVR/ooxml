/**
 * The space right before a collapsed caret, which an Enter at a soft word-wrap
 * boundary would strand at the end of the paragraph ("fox jumps" wrapped as
 * "fox " / "jumps" becoming "fox " and "jumps"). That invisible trailing
 * character still counts toward the line's width and can force an extra wrap,
 * so inline editors delete it before the browser splits the paragraph. Shared
 * by the React, Vue and Angular editors; each deletes the range itself, through
 * its collaboration `mutate` when the edit is connected.
 *
 * @module render/inline-trailing-space
 */

/** The range of the space before the caret, or `null` when there is none to drop. */
export function spaceBeforeCaretRange(doc: Document): Range | null {
	const selection = doc.getSelection();
	if (!selection || !selection.isCollapsed || selection.rangeCount === 0) {
		return null;
	}
	const { startContainer, startOffset } = selection.getRangeAt(0);
	if (startContainer.nodeType !== Node.TEXT_NODE || startOffset === 0) {
		return null;
	}
	if ((startContainer.textContent ?? '').charAt(startOffset - 1) !== ' ') {
		return null;
	}
	const range = doc.createRange();
	range.setStart(startContainer, startOffset - 1);
	range.setEnd(startContainer, startOffset);
	return range;
}
