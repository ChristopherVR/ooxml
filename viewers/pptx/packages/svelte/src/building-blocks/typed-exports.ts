import type { BuildingBlockToolbarGroupId } from 'ooxml-ui/pptx';
import type { Component } from 'svelte';

import ExternalContentBarComponent from '../viewer/components/ExternalContentBar.svelte';
import BuildingBlocksCanvasComponent from './BuildingBlocksCanvas.svelte';
import BuildingBlocksSlidesPaneComponent from './BuildingBlocksSlidesPane.svelte';
import BuildingBlocksToolbarComponent from './BuildingBlocksToolbar.svelte';
import type { ViewerBuildingBlocks } from './types';

/**
 * Explicitly typed public exports of the building-block components, for the same reason as
 * `viewer/components/typed-exports.ts`: the declaration bundling cannot resolve `.svelte`
 * specifiers, so the published `.d.ts` goes through annotated constants.
 */
export interface BuildingBlocksCanvasProps {
	blocks: ViewerBuildingBlocks;
	/** A viewer theme, as `PowerPointViewer`'s `theme`; unset inherits the host's tokens. */
	theme?: import('../viewer/types').ViewerTheme | undefined;
	class?: string;
}

export interface BuildingBlocksToolbarProps {
	blocks: ViewerBuildingBlocks;
	/** Show only these groups, in descriptor order. */
	include?: readonly BuildingBlockToolbarGroupId[] | undefined;
	class?: string;
}

export interface BuildingBlocksSlidesPaneProps {
	blocks: ViewerBuildingBlocks;
	class?: string;
}

/** PowerPoint's security bar; render it while `blocks.externalContent.barVisible`. */
export interface ExternalContentBarProps {
	/** Enable Content: `blocks.externalContent.enable`. */
	onenable: () => void;
	/** Close: `blocks.externalContent.close`. */
	onclose: () => void;
}

export const ExternalContentBar: Component<ExternalContentBarProps> =
	ExternalContentBarComponent as unknown as Component<ExternalContentBarProps>;

export const BuildingBlocksCanvas: Component<BuildingBlocksCanvasProps> =
	BuildingBlocksCanvasComponent as unknown as Component<BuildingBlocksCanvasProps>;

export const BuildingBlocksToolbar: Component<BuildingBlocksToolbarProps> =
	BuildingBlocksToolbarComponent as unknown as Component<BuildingBlocksToolbarProps>;

export const BuildingBlocksSlidesPane: Component<BuildingBlocksSlidesPaneProps> =
	BuildingBlocksSlidesPaneComponent as unknown as Component<BuildingBlocksSlidesPaneProps>;
