// @vitest-environment happy-dom
/**
 * The headless building-blocks API honours the host `googleFonts` policy the
 * way `PowerPointViewer` does: `false` (or an invalid mirror) locks Trust
 * Center > "Download fonts" off, a valid `{ baseUrl }` locks it on, and unset
 * leaves the store alone. Both host locks coexist.
 */
import type { GoogleFontsOption, ViewerOptionsStore } from 'ooxml-ui/pptx';
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

interface HarnessProps {
	googleFonts?: GoogleFontsOption;
	allowExternalContent?: boolean;
}

function Harness(props: HarnessProps) {
	useViewerBuildingBlocks({
		content: null,
		googleFonts: props.googleFonts,
		allowExternalContent: props.allowExternalContent,
	});
	return <div />;
}

async function render(props: HarnessProps) {
	await act(async () => root.render(<Harness {...props} />));
}

function fontState() {
	const store = captured.store!;
	return {
		locked: store.isLocked('trust', 'allowFontDownloads'),
		value: store.getOptions().trust.allowFontDownloads,
	};
}

describe('useViewerBuildingBlocks googleFonts', () => {
	afterEach(async () => {
		await act(async () => root?.unmount());
		container?.remove();
		localStorage.clear();
		captured.store = undefined;
	});

	async function mount(props: HarnessProps) {
		container = document.createElement('div');
		document.body.append(container);
		root = createRoot(container);
		await render(props);
	}

	it('locks Trust Center > Download fonts to the host value and follows changes', async () => {
		await mount({ googleFonts: false });
		expect(fontState()).toStrictEqual({ locked: true, value: false });
		await render({ googleFonts: { baseUrl: 'https://fonts.example.test/css2' } });
		expect(fontState()).toStrictEqual({ locked: true, value: true });
		await render({});
		expect(fontState()).toStrictEqual({ locked: false, value: true });
	});

	it('keeps both host locks together', async () => {
		await mount({ googleFonts: false, allowExternalContent: false });
		expect(fontState()).toStrictEqual({ locked: true, value: false });
		expect(captured.store!.isLocked('trust', 'allowExternalContent')).toBe(true);
		await render({ googleFonts: false });
		expect(fontState()).toStrictEqual({ locked: true, value: false });
		expect(captured.store!.isLocked('trust', 'allowExternalContent')).toBe(false);
	});
});
