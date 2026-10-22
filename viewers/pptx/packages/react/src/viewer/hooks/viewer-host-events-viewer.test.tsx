// @vitest-environment happy-dom
/**
 * The host events every binding reports the same way (issue 151, item 5), driven through a
 * mounted `PowerPointViewer`: `onChange` after a committed edit and `onPresentationChange` when
 * the slide show starts and stops.
 */
import { translationsEn } from 'ooxml-ui/pptx/i18n';
import React, { act, createRef } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

// oxlint-disable-next-line prefer-ending-with-an-expect
vi.mock<typeof import('react-i18next')>(import('react-i18next'), () => ({
	useTranslation: () => ({
		t: (key: string) => translationsEn[key] ?? key,
		i18n: {
			language: 'en',
			languages: ['en'],
			options: { resources: { en: {} } },
			changeLanguage: () => Promise.resolve(),
		},
	}),
}));

const { PptxHandler } = await import('pptx-viewer-core');
const { PowerPointViewer } = await import('../../index');
type ViewerHandle = import('../../index').PowerPointViewerHandle;

async function sampleDeck(): Promise<Uint8Array> {
	const { handler, data } = await PptxHandler.create({ initialSlideCount: 2 });
	try {
		return await handler.save(data.slides);
	} finally {
		handler.dispose();
	}
}

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	window.happyDOM?.setViewport({ width: 1600, height: 950 });
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => {
		root.unmount();
	});
	container.remove();
	globalThis.IS_REACT_ACT_ENVIRONMENT = false;
});

async function loaded(ref: React.RefObject<ViewerHandle | null>): Promise<void> {
	for (let attempt = 0; attempt < 80; attempt += 1) {
		if ((ref.current?.getSlideCount() ?? 0) > 0) {
			return;
		}
		await act(async () => {
			await new Promise((resolve) => {
				setTimeout(resolve, 20);
			});
		});
	}
	throw new Error('the viewer never finished loading its deck');
}

describe('host events through a mounted viewer', () => {
	it('fires onChange for a committed edit, not for the load', async () => {
		const onChange = vi.fn();
		const ref = createRef<ViewerHandle>();
		const content = await sampleDeck();
		await act(async () => {
			root.render(
				<PowerPointViewer
					ref={ref}
					content={content}
					filePath='events.pptx'
					canEdit
					onChange={onChange}
				/>,
			);
		});
		await loaded(ref);
		expect(onChange).not.toHaveBeenCalled();
		await act(async () => {
			(ref.current as ViewerHandle).addSlide();
		});
		expect(onChange).toHaveBeenCalled();
	});

	it('fires onPresentationChange when the slide show starts and stops', async () => {
		const onPresentationChange = vi.fn();
		const ref = createRef<ViewerHandle>();
		const content = await sampleDeck();
		await act(async () => {
			root.render(
				<PowerPointViewer
					ref={ref}
					content={content}
					filePath='events.pptx'
					canEdit
					onPresentationChange={onPresentationChange}
				/>,
			);
		});
		await loaded(ref);
		expect(onPresentationChange).not.toHaveBeenCalled();
		await act(async () => {
			(ref.current as ViewerHandle).setMode('present');
		});
		expect(onPresentationChange).toHaveBeenLastCalledWith(true);
		await act(async () => {
			(ref.current as ViewerHandle).setMode('edit');
		});
		expect(onPresentationChange).toHaveBeenLastCalledWith(false);
	});
});
