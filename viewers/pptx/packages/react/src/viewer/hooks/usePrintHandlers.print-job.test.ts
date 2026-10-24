// @vitest-environment jsdom
//
// Drives the hook the way `useIOHandlers` wires it, against a stubbed print
// window. jsdom (not happy-dom) so DOMPurify keeps every sibling page of the
// sanitised print body.
import { DEFAULT_VIEWER_OPTIONS } from 'ooxml-ui/pptx';
import type { ViewerOptions } from 'ooxml-ui/pptx';
import type { PptxData, PptxSlide } from 'pptx-viewer-core';
import { act, createElement } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import type { PrintSettings } from '../components/print-dialog-types';
import { ViewerOptionsContext } from '../components/viewer-options-context';
import { renderElementToTiledCanvas } from '../utils/export-helpers';
import { usePrintHandlers } from './usePrintHandlers';
import type { PrintHandlersResult, UsePrintHandlersInput } from './usePrintHandlers';

vi.mock(import('../utils/export-helpers'), async (importOriginal) => ({
	...(await importOriginal()),
	renderElementToTiledCanvas: vi.fn(async () => ({
		kind: 'canvas' as const,
		canvas: { toDataURL: () => 'data:image/png;base64,AAAA' } as unknown as HTMLCanvasElement,
		strategy: 'foreignObject' as const,
		width: 960,
		height: 540,
		tiled: false,
	})),
	waitForRender: vi.fn(async () => {}),
}));

// A German viewer: the print window title must come from the translator.
vi.mock(import('../i18n/use-translation'), async () => {
	const { translationsDe } = await import('pptx-viewer-locales');
	const t = (key: string, params?: unknown): string => {
		let message: string = translationsDe[key] ?? key;
		for (const [name, value] of Object.entries((params ?? {}) as Record<string, unknown>)) {
			message = message.replaceAll(`{{${name}}}`, String(value));
		}
		return message;
	};
	return { useTranslation: () => ({ t, locale: 'de' }) };
});

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

function makeSlides(count: number, hidden: number[] = []): PptxSlide[] {
	return Array.from({ length: count }, (_, i) => ({
		id: `s${i}`,
		rId: `rId${i}`,
		slideNumber: i + 1,
		hidden: hidden.includes(i) || undefined,
		elements: [
			{ id: `t${i}`, type: 'text', x: 20, y: 20, width: 300, height: 50, text: `Title ${i + 1}` },
		],
		notes: `Notes ${i + 1}`,
	})) as PptxSlide[];
}

function settings(overrides: Partial<PrintSettings> = {}): PrintSettings {
	return {
		printWhat: 'slides',
		orientation: 'landscape',
		colorMode: 'color',
		frameSlides: false,
		slidesPerPage: 6,
		slideRange: 'all',
		customRangeFrom: 1,
		customRangeTo: 1,
		...overrides,
	};
}

const write = vi.fn<(html: string) => void>();

beforeEach(() => {
	vi.mocked(renderElementToTiledCanvas).mockClear();
	write.mockClear();
	vi.spyOn(window, 'open').mockReturnValue({
		document: { open: vi.fn(), write, close: vi.fn() },
		focus: vi.fn(),
		print: vi.fn(),
		close: vi.fn(),
	} as unknown as Window);
});

let unmount: (() => void) | undefined;

afterEach(() => {
	unmount?.();
	unmount = undefined;
	vi.restoreAllMocks();
});

function renderPrintHandlers(
	input: Partial<UsePrintHandlersInput> & Pick<UsePrintHandlersInput, 'slides'>,
	options: ViewerOptions = DEFAULT_VIEWER_OPTIONS,
): () => PrintHandlersResult {
	let latest: PrintHandlersResult | undefined;
	const full: UsePrintHandlersInput = {
		activeSlideIndex: 0,
		canvasStageRef: { current: document.createElement('div') },
		setActiveSlideIndex: vi.fn(),
		canvasSize: { width: 960, height: 540 },
		...input,
	};
	function Harness(): null {
		latest = usePrintHandlers(full);
		return null;
	}
	const root = createRoot(document.createElement('div'));
	act(() => {
		root.render(
			createElement(ViewerOptionsContext.Provider, { value: options }, createElement(Harness)),
		);
	});
	unmount = () => act(() => root.unmount());
	return () => {
		if (!latest) {
			throw new Error('the hook did not render');
		}
		return latest;
	};
}

async function print(handlers: () => PrintHandlersResult, value: PrintSettings): Promise<void> {
	await act(async () => {
		await handlers().handlePrintSvg(value);
	});
}

/** The last document written into the print window. */
function printedHtml(): string {
	return String(write.mock.calls.at(-1)?.[0] ?? '');
}

describe('usePrintHandlers print jobs', () => {
	it('prints full-page slides as vector SVG with only the slide size wired in', async () => {
		const handlers = renderPrintHandlers({ slides: makeSlides(3) });

		await print(handlers, settings());

		expect(renderElementToTiledCanvas).not.toHaveBeenCalled();
		expect(window.open).toHaveBeenCalledOnce();
		const html = printedHtml();
		expect(html.match(/class="print-slide-page"/gu)).toHaveLength(3);
		expect(html).toContain('<svg');
		expect(html).not.toContain('data:image/png');
	});

	it('captures only the planned slide for notes pages', async () => {
		const setActiveSlideIndex = vi.fn();
		const handlers = renderPrintHandlers({
			slides: makeSlides(3),
			activeSlideIndex: 1,
			setActiveSlideIndex,
		});

		await print(handlers, settings({ printWhat: 'notes', slideRange: 'current' }));

		expect(renderElementToTiledCanvas).toHaveBeenCalledOnce();
		// Shows slide 2 for its capture, then restores the slide that was on screen.
		expect(setActiveSlideIndex.mock.calls).toStrictEqual([[1], [1]]);
		const html = printedHtml();
		expect(html).toContain('Notes 2');
		expect(html).not.toContain('Notes 1');
		expect(html).not.toContain('Notes 3');
	});

	it('prints hidden slides on the vector path when Print hidden slides is on', async () => {
		const slides = makeSlides(3, [1]);
		const options: ViewerOptions = {
			...DEFAULT_VIEWER_OPTIONS,
			advanced: { ...DEFAULT_VIEWER_OPTIONS.advanced, printHiddenSlides: true },
		};
		const pptxData = { slides, width: 960, height: 540 } as PptxData;
		const handlers = renderPrintHandlers({ slides, pptxData }, options);

		await print(handlers, settings());

		expect(renderElementToTiledCanvas).not.toHaveBeenCalled();
		expect(printedHtml().match(/class="print-slide-page"/gu)).toHaveLength(3);
	});

	it('leaves hidden slides out by default', async () => {
		const handlers = renderPrintHandlers({ slides: makeSlides(3, [1]) });

		await print(handlers, settings());

		expect(printedHtml().match(/class="print-slide-page"/gu)).toHaveLength(2);
	});

	it('prints the layout and master elements merged back into each vector slide', async () => {
		const slides = makeSlides(1);
		const handlers = renderPrintHandlers({
			slides,
			templateElementsBySlideId: {
				s0: [
					{
						id: 'master-logo',
						type: 'text',
						x: 0,
						y: 0,
						width: 400,
						height: 40,
						text: 'Brand',
					} as PptxSlide['elements'][number],
				],
			},
		});

		await print(handlers, settings());

		expect(printedHtml()).toContain('>Brand</tspan>');
	});

	it('titles the print window in the viewer language', async () => {
		const handlers = renderPrintHandlers({ slides: makeSlides(2) });

		await print(handlers, settings());
		expect(printedHtml()).toContain('<title>Folien</title>');

		await print(handlers, settings({ printWhat: 'outline' }));
		expect(printedHtml()).toContain('<title>Gliederung</title>');

		await print(handlers, settings({ printWhat: 'notes' }));
		expect(printedHtml()).toContain('<title>Notizseiten</title>');

		await print(handlers, settings({ printWhat: 'handouts', slidesPerPage: 4 }));
		expect(printedHtml()).toContain('<title>4 Folien pro Handzettel</title>');
	});

	it('opens nothing when the range selects only hidden slides', async () => {
		const handlers = renderPrintHandlers({ slides: makeSlides(2, [0]) });

		await print(handlers, settings({ slideRange: 'current' }));

		expect(window.open).not.toHaveBeenCalled();
		expect(renderElementToTiledCanvas).not.toHaveBeenCalled();
	});

	it('falls back to capturing the slides when the vector export throws', async () => {
		const slides = makeSlides(2);
		// A slide the SVG exporter cannot read (no element list) makes it throw.
		const broken = [{ ...slides[0], elements: undefined }, slides[1]] as unknown as PptxSlide[];
		const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
		const handlers = renderPrintHandlers({ slides: broken });

		await print(handlers, settings());

		expect(warn).toHaveBeenCalledWith(
			expect.stringContaining('falling back to raster'),
			expect.anything(),
		);
		expect(renderElementToTiledCanvas).toHaveBeenCalledTimes(2);
		expect(printedHtml()).toContain('data:image/png;base64,AAAA');
	});
});
