import type { Control, ControlGroup, SelectControl } from './controls';

/**
 * Draws control groups with plain DOM, for the bindings whose framework has no template of its own
 * (vanilla, Solid). Nodes are keyed by control id and updated in place, so a control keeps its
 * focus while the formatting state changes. Returns `update` and `destroy`.
 */
export function createControlsView(host: HTMLElement) {
	const nodes = new Map<string, HTMLElement>();
	const groups: HTMLElement[] = [];
	const keep = (select: HTMLSelectElement, control: SelectControl) => {
		const signature = control.options
			.map((option) => `${option.value}\u0000${option.label}`)
			.join('\u0001');
		if (select.dataset.signature !== signature) {
			select.dataset.signature = signature;
			select.replaceChildren(
				...control.options.map((option) => {
					const node = host.ownerDocument.createElement('option');
					node.value = option.value;
					node.textContent = option.label;
					return node;
				}),
			);
		}
		select.value = control.value;
	};
	const create = (control: Control): HTMLElement => {
		const doc = host.ownerDocument;
		if (control.kind === 'text') {
			const text = doc.createElement('span');
			text.setAttribute('role', 'status');
			return text;
		}
		if (control.kind === 'select') {
			const select = doc.createElement('select');
			select.addEventListener('change', () => {
				const current = nodes.get(control.id) as
					| (HTMLSelectElement & { control?: SelectControl })
					| undefined;
				current?.control?.change(select.value);
			});
			return select;
		}
		const button = doc.createElement('button');
		button.type = 'button';
		// Keep the document's selection while a button is pressed.
		button.addEventListener('mousedown', (event) => event.preventDefault());
		button.addEventListener('click', () => {
			const current = nodes.get(control.id) as (HTMLElement & { control?: Control }) | undefined;
			const latest = current?.control;
			if (latest && (latest.kind === 'toggle' || latest.kind === 'action')) latest.run();
		});
		return button;
	};
	const apply = (node: HTMLElement, control: Control) => {
		(node as HTMLElement & { control?: Control }).control = control;
		node.dataset.control = control.id;
		if (control.kind === 'text') {
			node.textContent = control.text;
			return;
		}
		node.setAttribute('aria-label', control.label);
		(node as HTMLButtonElement | HTMLSelectElement).disabled = control.disabled;
		if (control.kind === 'select') return keep(node as HTMLSelectElement, control);
		node.textContent = control.label;
		if (control.kind === 'toggle') node.setAttribute('aria-pressed', String(control.pressed));
		else node.removeAttribute('aria-pressed');
	};
	return {
		update(next: readonly ControlGroup[]) {
			const seen = new Set<string>();
			next.forEach((controls, index) => {
				let group = groups[index];
				if (!group) {
					group = host.ownerDocument.createElement('span');
					group.setAttribute('role', 'group');
					host.append(group);
					groups[index] = group;
				}
				const children: HTMLElement[] = [];
				for (const control of controls) {
					seen.add(control.id);
					let node = nodes.get(control.id);
					if (!node) {
						node = create(control);
						nodes.set(control.id, node);
					}
					apply(node, control);
					children.push(node);
				}
				// Re-append only when the order differs, so unchanged nodes are never detached.
				if (
					children.some((node, at) => group.children[at] !== node) ||
					group.children.length !== children.length
				)
					group.replaceChildren(...children);
			});
			for (const [id, node] of nodes)
				if (!seen.has(id)) {
					node.remove();
					nodes.delete(id);
				}
			while (groups.length > next.length) groups.pop()?.remove();
		},
		destroy() {
			host.replaceChildren();
			nodes.clear();
			groups.length = 0;
		},
	};
}
