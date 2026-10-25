// @vitest-environment happy-dom
/**
 * The headless building-blocks API honours the host `allowExternalContent`
 * switch the way `PowerPointViewer` does: a set value locks Trust Center >
 * "Allow external content" in its File > Options store, and unset leaves the
 * store alone.
 */
import type { ViewerOptionsStore } from 'ooxml-ui/pptx';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { useViewerBuildingBlocks } from './useViewerBuildingBlocks';

(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT =
	true;

const captured = vi.hoisted(() => ({ store: undefined as ViewerOptionsStore | undefined }));
vi.mock('./useViewerOptions', async (importOriginal) => {
	const actual = await importOriginal<typeof import('./useViewerOptions')>();
	return {
		...actual,
		useViewerOptions: () => {
			const result = actual.useViewerOptions();
			captured.store = result.optionsStore;
			return result;
		},
	};
});

let root: Root;
let container: HTMLDivElement;

function Harness(props: { allowExternalContent?: boolean }) {
	useViewerBuildingBlocks({ content: null, allowExternalContent: props.allowExternalContent });
	return <div />;
}

async function render(allowExternalContent?: boolean) {
	await act(async () => root.render(<Harness allowExternalContent={allowExternalContent} />));
}

function trustState() {
	const store = captured.store!;
	return {
		locked: store.isLocked('trust', 'allowExternalContent'),
		value: store.getOptions().trust.allowExternalContent,
	};
}

describe('useViewerBuildingBlocks allowExternalContent', () => {
	afterEach(async () => {
		await act(async () => root?.unmount());
		container?.remove();
		localStorage.clear();
		captured.store = undefined;
	});

	async function mount(allowExternalContent?: boolean) {
		container = document.createElement('div');
		document.body.append(container);
		root = createRoot(container);
		await render(allowExternalContent);
	}

	it('locks the Trust Center option to the host value and follows changes', async () => {
		await mount(false);
		expect(trustState()).toStrictEqual({ locked: true, value: false });
		await render(true);
		expect(trustState()).toStrictEqual({ locked: true, value: true });
		await render(undefined);
		expect(trustState()).toStrictEqual({ locked: false, value: true });
	});

	it('leaves the options store alone while unset', async () => {
		await mount(undefined);
		const setConstraints = vi.spyOn(captured.store!, 'setConstraints');
		await render(undefined);
		expect(setConstraints).not.toHaveBeenCalled();
		expect(trustState()).toStrictEqual({ locked: false, value: true });
	});
});
