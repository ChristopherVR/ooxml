import type { PptxHandler } from 'pptx-viewer-core';

import type { BuildingBlocksContractHarness } from './building-blocks-contract.mjs';

export const LINKED_CONTENT_ORIGIN: string;
export const LINKED_FILL_URL: string;
export function modelHasLinkedFill(model: unknown): boolean;
export function drawsLinkedContent(root: Element): boolean;

/** `JSZip` from `jszip`, structurally (each binding passes its own copy). */
export interface ExternalContentContractZip {
	loadAsync(data: Uint8Array): Promise<{
		file(path: string): { async(type: 'string'): Promise<string> } | null;
		file(path: string, data: string): unknown;
		generateAsync(options: { type: 'uint8array' }): Promise<Uint8Array>;
	}>;
}

/**
 * A building-blocks harness whose `mount` renders the canvas and the binding's
 * `ExternalContentBar` into the host element it is given.
 */
export interface ExternalContentContractHarness extends BuildingBlocksContractHarness {
	/** The blocks' `externalContent` (`ViewerExternalContentBar` from `ooxml-ui/pptx`), read now. */
	externalContent(): { barVisible: boolean; allowed: boolean; enable(): void; close(): void };
	/** The blocks' deck model now (`PptxSlide[]`), which keeps the link whatever is drawn. */
	slides(): readonly unknown[];
}

export function linkedContentDeck(
	Handler: typeof PptxHandler,
	JSZip: ExternalContentContractZip,
): Promise<Uint8Array>;
export function runExternalContentContract(input: {
	harness: ExternalContentContractHarness;
	expect: (actual: unknown, message?: string) => unknown;
	Handler: typeof PptxHandler;
	JSZip: unknown;
}): Promise<void>;
