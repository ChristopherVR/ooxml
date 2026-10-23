import type { PptxController } from 'ooxml-ui/pptx';
import { useSyncExternalStore } from 'react';

/**
 * The React store adapter for the shared controller convention (`getSnapshot`
 * / `subscribe` / typed actions, see `render/controller` in `ooxml-ui/pptx`).
 * One adapter serves every controller: it reads the snapshot through
 * `useSyncExternalStore`, so a render never tears and a controller change
 * re-renders exactly the components that read it.
 */
export function useController<S>(controller: PptxController<S>): S {
	return useSyncExternalStore(controller.subscribe, controller.getSnapshot, controller.getSnapshot);
}
