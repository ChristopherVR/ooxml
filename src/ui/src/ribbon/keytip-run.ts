/**
 * The second level of Office KeyTips for products that derive their tips from a ribbon panel: give
 * each visible command a key (the product's own Office keys where the command exists, otherwise two
 * letters of its label, never the prefix of another), show a badge over each and run the one whose
 * letters are typed. Word and Excel use the same loop and differ only in the key table and in how a
 * control is found and operated, so those stay with the product.
 */
export interface KeyTipTarget {
	/** The letters or digits to type; empty before {@link assignKeyTips}. */
	key: string;
	/** The element the badge sits under. */
	element: HTMLElement;
	/** What the tip does: open the control, or click it. */
	activate(): void;
}

const clash = (used: readonly string[], key: string) =>
	used.some((other) => other.startsWith(key) || key.startsWith(other));

/**
 * Makes every key unique and unambiguous. A key that is already set stays unless an earlier one
 * clashes with it (Office's own keys differ per tab, so two commands can share one); the others get
 * two letters from `label(target)`, then a number.
 */
export function assignKeyTips(
	targets: readonly KeyTipTarget[],
	label: (target: KeyTipTarget) => string,
): void {
	const used: string[] = [];
	for (const target of targets) {
		if (target.key && !clash(used, target.key)) used.push(target.key);
		else target.key = '';
	}
	for (const target of targets.filter((item) => !item.key)) {
		const source =
			label(target)
				.toUpperCase()
				.replace(/[^A-Z0-9]/g, '') || 'ZZ';
		let key = '';
		for (let a = 0; a < source.length && !key; a++)
			for (let b = a + 1; b < source.length && !key; b++) {
				const candidate = source[a]! + source[b]!;
				if (!clash(used, candidate)) key = candidate;
			}
		for (let n = 10; !key; n++) if (!clash(used, String(n))) key = String(n);
		used.push(key);
		target.key = key;
	}
}

/**
 * Shows fixed badges over `targets` (in `root`, so product CSS styles them through `className`) and
 * runs the one whose letters are typed. Escape, a pointer press or a key with no match ends it;
 * `onDone` runs after any end. Returns a function that stops it.
 */
export function runKeyTips(
	root: HTMLElement,
	targets: readonly KeyTipTarget[],
	onDone: () => void,
	className: string,
): () => void {
	const doc = root.ownerDocument;
	const shown = targets.map((target) => {
		const node = doc.createElement('span');
		node.className = className;
		node.textContent = target.key;
		node.setAttribute('aria-hidden', 'true');
		const box = target.element.getBoundingClientRect();
		node.style.position = 'fixed';
		node.style.left = `${box.left + box.width / 2}px`;
		node.style.top = `${box.bottom - 10}px`;
		root.append(node);
		return { node, target };
	});
	let typed = '';
	const stop = () => {
		for (const { node } of shown) node.remove();
		root.removeEventListener('keydown', onKey, true);
		root.removeEventListener('pointerdown', stop, true);
		onDone();
	};
	function onKey(event: KeyboardEvent) {
		if (event.key === 'Alt' || event.key === 'Shift') return;
		event.preventDefault();
		event.stopPropagation();
		if (event.key === 'Escape' || event.ctrlKey || event.metaKey) return stop();
		typed += event.key.toUpperCase();
		const matches = shown.filter(({ target }) => target.key.startsWith(typed));
		if (!matches.length) return stop();
		for (const { node, target } of shown) node.hidden = !target.key.startsWith(typed);
		const exact = matches.find(({ target }) => target.key === typed);
		if (exact) {
			stop();
			exact.target.activate();
		}
	}
	root.addEventListener('keydown', onKey, true);
	root.addEventListener('pointerdown', stop, true);
	return stop;
}
