// Large ribbon buttons caption their label on at most two lines, as Excel does: the break goes
// at the space that best balances the two lines, words are never cut and nothing is truncated.
// The group then grows to the widest line.

/** The label with a `\n` at the most balanced space (one line for a single word). */
export function balancedCaption(label: string): string {
	const text = label.trim().replace(/\s+/g, ' ');
	let best = -1;
	let widest = text.length;
	for (let i = text.indexOf(' '); i >= 0; i = text.indexOf(' ', i + 1)) {
		const longer = Math.max(i, text.length - i - 1);
		if (longer < widest) {
			widest = longer;
			best = i;
		}
	}
	return best < 0 ? text : `${text.slice(0, best)}\n${text.slice(best + 1)}`;
}

/** Writes a large button's caption (two balanced lines). */
export function setLargeCaption(caption: HTMLElement, label: string): void {
	caption.textContent = balancedCaption(label);
	caption.classList.add('ribbon-caption');
}
