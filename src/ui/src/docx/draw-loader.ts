// The editor's link to its Draw tab (`draw-tab.ts`), which loads on first use.
import type { EditorCore } from './editor-core';

type DrawTab = typeof import('./draw-tab');
let loaded: DrawTab | undefined;
let loading: Promise<DrawTab> | undefined;

/** The Draw tab's module once it has loaded. */
export const drawTab = (): DrawTab | undefined => loaded;

/** Loads the Draw tab, builds its panel in `core`'s ribbon and refreshes the editor's controls. */
export function loadDrawTab(core: EditorCore): Promise<DrawTab> {
	loading ??= import('./draw-tab').then((module) => (loaded = module));
	return loading.then((module) => {
		module.mountDrawPanel(core);
		core.refreshControls();
		return module;
	});
}
