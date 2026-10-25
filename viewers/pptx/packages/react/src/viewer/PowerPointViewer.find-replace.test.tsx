// @vitest-environment happy-dom
/**
 * The find bar through the whole viewer. The viewer clears the selection after every slide
 * change, in an effect that runs after the render in which Find set both the slide and the
 * selection, so a match on another slide was selected and then unselected straight away. A slide
 * change of the user's own must still clear the selection.
 */
import { translationsEn } from 'ooxml-ui/pptx/i18n';
import { PptxHandler } from 'pptx-viewer-core';
import React, { act, createRef } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, expect, test, vi } from 'vitest';

vi.mock(import('react-i18next'), () => ({
	useTranslation: () => ({
		t: (key: string, params?: Record<string, unknown>) =>
			(translationsEn[key] ?? key).replace(/\{\{(\w+)\}\}/gu, (_, name: string) =>
				String(params?.[name] ?? ''),
			),
		i18n: {
			language: 'en',
			languages: ['en'],
			options: { resources: { en: {} } },
			changeLanguage: () => Promise.resolve(),
		},
	}),
}));
const { PowerPointViewer } = await import('../index');
type Handle = import('../index').PowerPointViewerHandle;

let root: Root | undefined;
let container: HTMLDivElement | undefined;
const hadActEnvironment = Object.hasOwn(globalThis, 'IS_REACT_ACT_ENVIRONMENT');
const previousActEnvironment = globalThis.IS_REACT_ACT_ENVIRONMENT;

afterEach(async () => {
	await act(async () => root?.unmount());
	container?.remove();
	root = undefined;
	container = undefined;
	if (hadActEnvironment) {
		globalThis.IS_REACT_ACT_ENVIRONMENT = previousActEnvironment;
	} else {
		delete globalThis.IS_REACT_ACT_ENVIRONMENT;
	}
});

/** Two slides with one "quarter" match each. */
async function deck(): Promise<Uint8Array> {
	const { handler, data } = await PptxHandler.create({ initialSlideCount: 2 });
	data.slides.forEach((slide, index) => {
		const value = index ? 'quarter two' : 'quarter one';
		slide.elements = [
			{ id: `t${index}`, type: 'text', x: 67, y: 40, width: 400, height: 60, text: value },
		];
		slide.isDirty = true;
	});
	const bytes = await handler.save(data.slides);
	handler.dispose();
	return bytes;
}

/** The viewer's handle, read fresh: each render rebuilds it over that render's state. */
async function mountViewer(): Promise<() => Handle> {
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	const content = await deck();
	container = document.createElement('div');
	document.body.append(container);
	root = createRoot(container);
	const ref = createRef<Handle>();
	await act(async () => {
		root!.render(<PowerPointViewer ref={ref} content={content} canEdit />);
	});
	await vi.waitFor(
		async () => {
			await act(async () => {
				await new Promise((done) => {
					setTimeout(done, 10);
				});
			});
			expect(ref.current?.getSlideCount()).toBe(2);
		},
		{ timeout: 15000 },
	);
	return () => ref.current!;
}

/** The id of the element on `slide` whose text is `value`. */
const idOf = (viewer: () => Handle, slide: number, value: string) => {
	const elements = viewer().getSlides()[slide]?.elements ?? [];
	return elements.find((e) => (e as { text?: string }).text === value)?.id;
};

async function type(input: HTMLInputElement, value: string): Promise<void> {
	await act(async () => {
		Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!.call(input, value);
		input.dispatchEvent(new Event('input', { bubbles: true }));
	});
}

test('next selects the match on another slide, and a slide change of its own still clears', async () => {
	const viewer = await mountViewer();
	await act(async () => {
		window.dispatchEvent(new KeyboardEvent('keydown', { key: 'h', ctrlKey: true, bubbles: true }));
	});
	const query = container!.querySelector<HTMLInputElement>('input[aria-label="Search text"]')!;
	expect(query).not.toBeNull();
	await type(query, 'quarter');
	expect(viewer().getActiveSlideIndex()).toBe(0);
	expect(viewer().getSelectedElementIds()).toStrictEqual([idOf(viewer, 0, 'quarter one')]);

	await act(async () => {
		container!.querySelector<HTMLButtonElement>('button[aria-label="Next match"]')!.click();
	});
	expect(viewer().getActiveSlideIndex()).toBe(1);
	expect(viewer().getSelectedElementIds()).toStrictEqual([idOf(viewer, 1, 'quarter two')]);

	await act(async () => viewer().setActiveSlideIndex(0));
	expect(viewer().getSelectedElementIds()).toStrictEqual([]);
	await act(async () => viewer().selectElements([idOf(viewer, 0, 'quarter one')!]));
	await act(async () => viewer().setActiveSlideIndex(1));
	expect(viewer().getSelectedElementIds()).toStrictEqual([]);
});
