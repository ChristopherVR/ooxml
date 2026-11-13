import type { PptxHandler } from 'pptx-viewer-core';

export interface BuildingBlocksContractOptions {
	content: Uint8Array;
	canEdit: boolean;
	onRequest: (request: unknown) => void;
	onDirtyChange: (dirty: boolean) => void;
	/** A host-owned room (`externalSession`) the contract joins and leaves through the options. */
	collaboration?: unknown;
	googleFonts?: false | { baseUrl: string };
	allowExternalContent?: boolean;
}

/** `CollaborationShellState` from `ooxml-ui/pptx`, structurally. */
export interface BuildingBlocksContractShellState {
	canEdit: boolean;
	status: string;
	remoteUsers: readonly unknown[];
	connectedCount: number;
}

/** `CollaborationState` from `ooxml-ui/pptx`, structurally. */
export interface BuildingBlocksContractCollaborationState {
	active: boolean;
	roomId: string | undefined;
	status: string;
	synced: boolean;
	people: readonly { clientId: number; name: string; self: boolean }[];
}

/** What a binding supplies to run the contract against its building blocks. */
export interface BuildingBlocksContractHarness {
	readonly name: string;
	/**
	 * Whether the binding renders the compact toolbar (`describeBuildingBlockToolbar`) and marks its
	 * blocks with `data-pptx-block="canvas" | "toolbar" | "slides-pane"`. React's full ribbon does not.
	 */
	readonly toolbar: boolean;
	/** Mount the canvas, toolbar and slides pane into `host` with `options`. */
	mount(host: HTMLElement, options: BuildingBlocksContractOptions): Promise<void>;
	/** Apply new options the way a host does (a prop change, `update`, a signal). */
	update(options: BuildingBlocksContractOptions): Promise<void>;
	/** The current `actions` object (`ViewerBuildingBlockActions` from `ooxml-ui/pptx`). */
	actions(): unknown;
	state(): {
		mode: string;
		loading: boolean;
		error: string | null;
		slideCount: number;
		shellState: BuildingBlocksContractShellState;
		collaborationState: BuildingBlocksContractCollaborationState;
	};
	/** Press the toolbar's Pictures control. */
	clickPictures(): Promise<void>;
	/** Let the binding flush pending updates (React `act`, Angular change detection). */
	settle(): Promise<void>;
	destroy(): void;
}

export const COMPACT_TOOLBAR_CONTROLS: readonly string[];
export const BUILDING_BLOCK_GROUP_IDS: readonly string[];
export function buildingBlocksContractDeck(Handler: typeof PptxHandler): Promise<Uint8Array>;
export function runBuildingBlocksContract(input: {
	harness: BuildingBlocksContractHarness;
	/** The caller's Vitest `expect` and `vi` (kept structural so no Vitest copy is pinned here). */
	expect: (actual: unknown, message?: string) => unknown;
	vi: { fn: () => unknown };
	Handler: typeof PptxHandler;
	/** The binding's own `yjs` (the room's document must be the binding's `Y.Doc` class). */
	Y: { Doc: new () => { clientID: number; destroy(): void } };
}): Promise<void>;
