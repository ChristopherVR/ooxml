import { createSaveRevisionController } from 'ooxml-ui/pptx';
import type { SaveRevisionController, SaveRevisionControllerDeps } from 'ooxml-ui/pptx';
import { useRef } from 'react';

/**
 * The per-viewer save revision tracker behind `getSaveRevision`, `getSaveSnapshot` and
 * `acknowledgeSave` (see `createSaveRevisionController` in `ooxml-ui/pptx`). The controller is
 * built once; `deps` is read through a ref that is refreshed on every render, so the change
 * sources, serializer and clean-up callbacks are always the latest render's.
 */
export function useSaveRevision(deps: SaveRevisionControllerDeps): SaveRevisionController {
	const depsRef = useRef(deps);
	depsRef.current = deps;
	const controllerRef = useRef<SaveRevisionController | null>(null);
	controllerRef.current ??= createSaveRevisionController({
		getSources: () => depsRef.current.getSources(),
		serialize: () => depsRef.current.serialize(),
		markClean: () => depsRef.current.markClean(),
		discardRecovery: () => depsRef.current.discardRecovery(),
	});
	return controllerRef.current;
}
