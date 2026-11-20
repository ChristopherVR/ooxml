/**
 * WCAG 2.5.3 (Label in Name) for unit tests: a control's accessible name must contain the text it
 * shows. The comparison follows axe's `label-content-name-mismatch`, which the accessibility gates
 * run in a browser: words in order, ignoring case, punctuation, symbols and parentheticals, with
 * words split the way `Intl.Segmenter` splits them (so Chinese compares word by word). A shown text
 * of a single letter, or the font samples `Aa` and `Abc`, is not a label.
 */

const SEGMENTER = new Intl.Segmenter(undefined, { granularity: 'word' });

const words = (text: string): string[] =>
	[
		...SEGMENTER.segment(
			text
				.normalize('NFKD')
				.replace(/\([^()]*\)/gu, '')
				.toLowerCase()
				.replace(/[^\p{L}\p{N}]+/gu, ' '),
		),
	]
		.map((part) => part.segment.trim())
		.filter(Boolean);

/** Whether `name` contains every word of `label`, contiguously and in order. */
export function labelInName(label: string, name: string): boolean {
	const needle = words(label);
	if (needle.length === 0 || (needle.length === 1 && /^(?:\p{L}|aa|abc)$/u.test(needle[0] ?? '')))
		return true;
	const haystack = words(name);
	for (let start = 0; start + needle.length <= haystack.length; start++)
		if (needle.every((word, at) => haystack[start + at] === word)) return true;
	return false;
}

const NAMED_FROM_CONTENT = [
	'button',
	'a[href]',
	'[role="button"]',
	'[role="tab"]',
	'[role="menuitem"]',
	'[role="link"]',
	'[role="option"]',
	'[role="treeitem"]',
].join(', ');

/** The text a control shows: its text nodes, without those of hidden descendants. */
function shownText(control: Element): string {
	let text = '';
	for (const node of control.childNodes) {
		if (node.nodeType === 3) text += node.textContent ?? '';
		else if (node instanceof Element && !node.hasAttribute('hidden')) text += shownText(node);
	}
	return text;
}

/**
 * Each control under `root` (and in the open shadow roots below it) whose `aria-label` leaves out
 * the text it shows, as `"shown" -> "name"`.
 */
export function labelNameMismatches(root: ParentNode): string[] {
	const found: string[] = [];
	const visit = (scope: ParentNode) => {
		for (const control of scope.querySelectorAll<HTMLElement>(NAMED_FROM_CONTENT)) {
			const name = control.getAttribute('aria-label');
			const shown = shownText(control).replace(/\s+/gu, ' ').trim();
			if (name && shown && !labelInName(shown, name)) found.push(`"${shown}" -> "${name}"`);
		}
		for (const element of scope.querySelectorAll('*'))
			if (element.shadowRoot) visit(element.shadowRoot);
	};
	visit(root);
	return found;
}
