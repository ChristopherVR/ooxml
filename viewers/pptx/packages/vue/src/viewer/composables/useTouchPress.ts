import { installTouchPress } from 'ooxml-ui/pptx';
import { watchEffect } from 'vue';
import type { Ref } from 'vue';

/**
 * Turns a touch long-press into `contextmenu` and a double-tap into `dblclick`
 * on the controls under `root` that opt in with `data-pptx-touch-press`
 * (issue #290: iOS Safari fires neither). Re-installs when the element changes.
 */
export function useTouchPress(root: Ref<HTMLElement | null>): void {
	watchEffect((onCleanup) => {
		const element = root.value;
		if (element) {
			onCleanup(installTouchPress(element));
		}
	});
}
