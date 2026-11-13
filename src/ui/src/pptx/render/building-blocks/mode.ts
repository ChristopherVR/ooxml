/**
 * The `mode` every binding's building blocks report: the viewer's mode, except that a deck the
 * blocks may not edit reads `preview`. The canvas keeps its own mode (React's internal `edit`
 * with `canEdit: false`); hosts see the effective one, so the same deck and permission report the
 * same mode in every binding.
 *
 * @module render/building-blocks/mode
 */
import type { ViewerMode } from '../../types';

export function resolveBuildingBlockMode(mode: ViewerMode, canEdit: boolean): ViewerMode {
	return mode === 'edit' && !canEdit ? 'preview' : mode;
}
