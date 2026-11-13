import type { PptxHandler } from 'pptx-viewer-core';

import type { BuildingBlocksContractHarness } from './building-blocks-contract.mjs';

export function elementRenderErrorContractDeck(Handler: typeof PptxHandler): Promise<Uint8Array>;
/**
 * Run the element render error contract against a building-blocks harness whose `mount`
 * forwards every option it is given, `onElementRenderError` included.
 */
export function runElementRenderErrorContract(input: {
	harness: BuildingBlocksContractHarness;
	expect: (actual: unknown, message?: string) => unknown;
	vi: { fn: (...args: never[]) => unknown; spyOn: (...args: never[]) => unknown };
	Handler: typeof PptxHandler;
}): Promise<void>;
