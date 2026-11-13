// @vitest-environment happy-dom
/**
 * React's building blocks (`SlideCanvas` and `SlidesPane` spread with the hook's props) against the
 * element render error contract every binding passes
 * (`packages/shared/test-utils/element-render-error-contract.mjs`).
 */
import { PptxHandler } from 'pptx-viewer-core';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { expect, test, vi } from 'vitest';

import type {
	BuildingBlocksContractHarness,
	BuildingBlocksContractOptions,
} from '../../../../shared/test-utils/building-blocks-contract.mjs';
import { runElementRenderErrorContract } from '../../../../shared/test-utils/element-render-error-contract.mjs';
import { SlideCanvas } from '../components/SlideCanvas';
import { SlidesPaneSidebar } from '../components/SlidesPaneSidebar';
import { useViewerBuildingBlocks } from './useViewerBuildingBlocks';
import type { ViewerBuildingBlocksResult } from './useViewerBuildingBlocks';

(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT =
	true;

function reactHarness(): BuildingBlocksContractHarness {
	let root: Root | undefined;
	let latest: ViewerBuildingBlocksResult | undefined;
	function Harness({ options }: { options: BuildingBlocksContractOptions }) {
		latest = useViewerBuildingBlocks({ ...options, autosaveEnabled: false });
		return (
			<>
				<SlidesPaneSidebar {...latest.slidesPaneProps} />
				<SlideCanvas {...latest.canvasProps} />
			</>
		);
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
		clickPictures: async () => {},
		async settle() {
			await act(async () => {
				await new Promise<void>((resolve) => {
					setTimeout(resolve, 0);
				});
			});
		},
		destroy: () => {
			act(() => root?.unmount());
			root = undefined;
		},
	};
}

test('passes the shared element render error contract', async () => {
	await runElementRenderErrorContract({
		harness: reactHarness(),
		expect,
		vi,
		Handler: PptxHandler,
	});
}, 60000);
