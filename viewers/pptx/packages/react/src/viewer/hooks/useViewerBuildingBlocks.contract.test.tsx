// @vitest-environment happy-dom
/**
 * React's `useViewerBuildingBlocks` against the building-blocks parity contract every binding
 * passes (`packages/shared/test-utils/building-blocks-contract.mjs`).
 */
import { PptxHandler } from 'pptx-viewer-core';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { expect, test, vi } from 'vitest';

import { runBuildingBlocksContract } from '../../../../shared/test-utils/building-blocks-contract.mjs';
import type {
	BuildingBlocksContractHarness,
	BuildingBlocksContractOptions,
} from '../../../../shared/test-utils/building-blocks-contract.mjs';
import { useViewerBuildingBlocks } from './useViewerBuildingBlocks';
import type { ViewerBuildingBlocksResult } from './useViewerBuildingBlocks';

(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT =
	true;

function reactHarness(): BuildingBlocksContractHarness {
	let root: Root | undefined;
	let latest: ViewerBuildingBlocksResult | undefined;
	function Harness({ options }: { options: BuildingBlocksContractOptions }) {
		latest = useViewerBuildingBlocks({ ...options, autosaveEnabled: false });
		return <div />;
	}
	const render = async (options: BuildingBlocksContractOptions) => {
		await act(async () => root!.render(<Harness options={options} />));
	};
	return {
		name: 'react',
		toolbar: false,
		async mount(host, options) {
			root = createRoot(host);
			await render(options);
		},
		update: render,
		actions: () => latest!.actions,
		state: () => ({
			mode: latest!.mode,
			loading: latest!.loading,
			error: latest!.error,
			slideCount: latest!.slidesPaneProps.slides.length,
		}),
		async clickPictures() {
			await act(async () => latest!.toolbarProps.onOpenImagePicker?.());
		},
		async settle() {
			await act(async () => {
				await new Promise<void>((resolve) => {
					setTimeout(resolve, 0);
				});
			});
		},
		destroy: () => act(() => root?.unmount()),
	};
}

test('passes the shared building-blocks contract', async () => {
	await runBuildingBlocksContract({ harness: reactHarness(), expect, vi, Handler: PptxHandler });
}, 60000);
