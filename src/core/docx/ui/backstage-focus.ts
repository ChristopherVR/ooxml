/** Make the rest of the editor inert while the File view is open and restore its opening control. */
export function backstageFocus(element: HTMLElement) {
	let opener: HTMLElement | null = null;
	const siblings = new Map<HTMLElement, boolean>();
	const active = () => (element.getRootNode() as Document | ShadowRoot).activeElement;
	return {
		open() {
			const current = active();
			if (current instanceof HTMLElement && !element.contains(current)) opener = current;
			for (const sibling of element.parentElement?.children ?? []) {
				if (!(sibling instanceof HTMLElement) || sibling === element || siblings.has(sibling))
					continue;
				siblings.set(sibling, Boolean(sibling.inert));
				sibling.inert = true;
			}
		},
		close() {
			for (const [sibling, wasInert] of siblings) sibling.inert = wasInert;
			siblings.clear();
			if (opener?.isConnected) opener.focus();
			opener = null;
		},
	};
}
