import type {
	BuildingBlockToolbarControlId,
	BuildingBlockToolbarGroup,
	BuildingBlockToolbarGroupId,
} from 'ooxml-ui/pptx';

/** Replace children only when they differ, so a focused control is not moved by a refresh. */
function sync(parent: Element, children: readonly Node[]): void {
	const now = parent.childNodes;
	if (now.length === children.length && children.every((child, index) => now[index] === child)) {
		return;
	}
	parent.replaceChildren(...children);
}

/**
 * Keyed DOM for the compact toolbar (`describeBuildingBlockToolbar`): one `<button>` per control,
 * created once and updated in place, inside one `<span data-group>` per group.
 */
export function createToolbarView(
	doc: Document,
	run: (id: BuildingBlockToolbarControlId) => void,
): (
	host: HTMLElement,
	groups: readonly BuildingBlockToolbarGroup[],
	translate: (key: string) => string,
	include?: readonly BuildingBlockToolbarGroupId[],
) => void {
	const nodes = new Map<string, HTMLElement>();
	const node = <K extends 'button' | 'span'>(key: string, tag: K): HTMLElementTagNameMap[K] => {
		let found = nodes.get(key) as HTMLElementTagNameMap[K] | undefined;
		if (!found) {
			found = doc.createElement(tag);
			nodes.set(key, found);
		}
		return found;
	};
	return (host, groups, translate, include) => {
		const shown = include ? groups.filter((group) => include.includes(group.id)) : groups;
		sync(
			host,
			shown.map((group) => {
				const span = node(`group-${group.id}`, 'span');
				span.dataset.group = group.id;
				span.setAttribute('role', 'group');
				sync(
					span,
					group.controls.map((control) => {
						const button = node(control.id, 'button');
						button.type = 'button';
						button.dataset.control = control.id;
						button.textContent = translate(control.labelKey);
						button.disabled = control.disabled;
						if (control.pressed === undefined) {
							button.removeAttribute('aria-pressed');
						} else {
							button.setAttribute('aria-pressed', String(control.pressed));
						}
						button.onclick = () => run(control.id);
						return button;
					}),
				);
				return span;
			}),
		);
	};
}
