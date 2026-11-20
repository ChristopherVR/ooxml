/**
 * Dropdown open-state + outside-click dismissal: the Vue replacement for the
 * `useState` + `useEffect(mousedown)` pattern each React ribbon section repeats.
 *
 * Bind `root` to the dropdown's wrapper element (`ref="…"` /
 * `:ref="dd.root"`); while `open` is true a document `mousedown` outside `root`
 * closes it, and Escape closes it and gives focus back to the wrapper's
 * `aria-haspopup="menu"` button. Call once per independent dropdown in a section.
 */
import { listenForPopupMenuDismiss } from 'ooxml-ui/pptx';
import { onScopeDispose, ref, watch } from 'vue';

export interface Dropdown {
	/** Reactive open flag. */
	open: ReturnType<typeof ref<boolean>>;
	/** Template ref for the dropdown wrapper (outside-click boundary). */
	root: ReturnType<typeof ref<HTMLElement | null>>;
	toggle: () => void;
	close: () => void;
}

export function useDropdown(): Dropdown {
	const open = ref(false);
	const root = ref<HTMLElement | null>(null);

	function onDocMouseDown(e: MouseEvent): void {
		if (root.value && !root.value.contains(e.target as Node)) {
			open.value = false;
		}
	}

	let detachEscape: (() => void) | undefined;
	watch(open, (isOpen) => {
		detachEscape?.();
		detachEscape = undefined;
		if (isOpen) {
			document.addEventListener('mousedown', onDocMouseDown);
			detachEscape = listenForPopupMenuDismiss(document, {
				onDismiss: () => {
					open.value = false;
				},
				trigger: root.value?.querySelector<HTMLElement>('[aria-haspopup="menu"]'),
			});
		} else {
			document.removeEventListener('mousedown', onDocMouseDown);
		}
	});

	onScopeDispose(() => {
		document.removeEventListener('mousedown', onDocMouseDown);
		detachEscape?.();
	});

	return {
		open,
		root,
		toggle: () => {
			open.value = !open.value;
		},
		close: () => {
			open.value = false;
		},
	};
}
