/**
 * Tests for pure logic and type contract of useThemeHandlers (GAP-E3).
 *
 * The theme switching coordination logic delegates to PptxHandler.
 * These tests verify the handler result shape and edge-case handling
 * without mounting React components.
 */
import { PRESET_THEMES } from 'ooxml-ui/pptx';
import type { PptxElement, PptxHandler, PptxSlide } from 'pptx-viewer-core';
import { describe, it, expect, vi } from 'vitest';

import { useThemeHandlers } from './useThemeHandlers';
import type { ThemeHandlersResult, UseThemeHandlersInput } from './useThemeHandlers';

// ---------------------------------------------------------------------------
// Type-level assertions: ensure the new methods exist in the result type
// ---------------------------------------------------------------------------

describe('themeHandlersResult type contract', () => {
	it('should include handleGetAvailableThemes in the result interface', () => {
		// Compile-time check: the property must exist on ThemeHandlersResult
		const check: keyof ThemeHandlersResult = 'handleGetAvailableThemes';
		expect(check).toBe('handleGetAvailableThemes');
	});

	it('should include handleSwitchTheme in the result interface', () => {
		const check: keyof ThemeHandlersResult = 'handleSwitchTheme';
		expect(check).toBe('handleSwitchTheme');
	});

	it('should include all original handlers in the result interface', () => {
		const keys: Array<keyof ThemeHandlersResult> = [
			'handleApplyTheme',
			'handleUpdateThemeColorScheme',
			'handleUpdateThemeFontScheme',
			'handleUpdateThemeName',
			'handleApplyThemeToPresentation',
			'handleApplyThemeData',
			'handleSetTemplateBackground',
			'handleGetTemplateBackgroundColor',
			'handleGetAvailableThemes',
			'handleSwitchTheme',
		];
		expect(keys).toHaveLength(10);
	});
});

// ---------------------------------------------------------------------------
// Input type validation
// ---------------------------------------------------------------------------

describe('useThemeHandlersInput contract', () => {
	it('should accept all required input fields', () => {
		const input: UseThemeHandlersInput = {
			handlerRef: { current: null },
			serializeSlides: vi.fn<() => void>().mockResolvedValue(null),
			setContent: vi.fn<() => void>(),
			onContentChange: undefined,
			setTheme: vi.fn<() => void>(),
			setSlideMasters: vi.fn<() => void>(),
			slideMasters: [],
			history: {
				markDirty: vi.fn<() => void>(),
			} as unknown as UseThemeHandlersInput['history'],
			setSlides: vi.fn<() => void>(),
			templateElementsBySlideId: {},
			setTemplateElementsBySlideId: vi.fn<() => void>(),
			theme: undefined,
			bumpHistory: vi.fn<() => void>(),
		};

		expect(input.handlerRef.current).toBeNull();
		expect(input.slideMasters).toStrictEqual([]);
	});
});

// ---------------------------------------------------------------------------
// handleUpdateThemeColorScheme: template-elements layer re-colour
// ---------------------------------------------------------------------------

describe('handleUpdateThemeColorScheme', () => {
	it('re-colours templateElementsBySlideId alongside slides, using the same old/new maps', async () => {
		const OFFICE_ACCENT1 = '#4472C4';
		const ION_ACCENT1 = '#B01513';

		let capturedTemplateElements: Record<string, PptxElement[]> = {
			'layout-slide-1': [
				{
					type: 'shape',
					id: 'bg_1',
					x: 0,
					y: 0,
					width: 200,
					height: 100,
					shapeStyle: { fillColor: OFFICE_ACCENT1 },
				} as PptxElement,
			],
		};

		const input: UseThemeHandlersInput = {
			handlerRef: {
				current: {
					updateThemeColorScheme: vi.fn().mockResolvedValue(undefined),
				} as unknown as PptxHandler,
			},
			serializeSlides: vi.fn<() => void>().mockResolvedValue(null),
			setContent: vi.fn<() => void>(),
			onContentChange: undefined,
			setTheme: vi.fn<() => void>(),
			setSlideMasters: vi.fn<() => void>(),
			slideMasters: [],
			history: { markDirty: vi.fn<() => void>() } as unknown as UseThemeHandlersInput['history'],
			setSlides: vi.fn<() => void>(),
			templateElementsBySlideId: capturedTemplateElements,
			setTemplateElementsBySlideId: vi.fn((updater) => {
				capturedTemplateElements =
					typeof updater === 'function' ? updater(capturedTemplateElements) : updater;
			}),
			theme: {
				colorScheme: {
					dk1: '#000000',
					lt1: '#FFFFFF',
					dk2: '#44546A',
					lt2: '#E7E6E6',
					accent1: OFFICE_ACCENT1,
					accent2: '#ED7D31',
					accent3: '#A5A5A5',
					accent4: '#FFC000',
					accent5: '#5B9BD5',
					accent6: '#70AD47',
					hlink: '#0563C1',
					folHlink: '#954F72',
				},
			} as UseThemeHandlersInput['theme'],
			bumpHistory: vi.fn<() => void>(),
		};

		const { handleUpdateThemeColorScheme } = useThemeHandlers(input);
		await handleUpdateThemeColorScheme({
			...input.theme!.colorScheme!,
			accent1: ION_ACCENT1,
		});

		expect(input.setTemplateElementsBySlideId).toHaveBeenCalledWith(expect.any(Function));
		const patched = capturedTemplateElements['layout-slide-1']![0] as {
			shapeStyle?: { fillColor?: string };
		};
		expect(patched.shapeStyle?.fillColor).toBe(ION_ACCENT1);
	});
});

// ---------------------------------------------------------------------------
// Design > Themes and Fonts: applied in place, never by reloading the deck
// ---------------------------------------------------------------------------

describe('theme picks apply in place', () => {
	const FONTS = {
		majorFont: { latin: 'Calibri Light' },
		minorFont: { latin: 'Calibri' },
	} as NonNullable<UseThemeHandlersInput['theme']>['fontScheme'];

	function inPlaceInput(handler: Partial<PptxHandler>) {
		const slides = [
			{ id: 's1', elements: [] },
			{ id: 's2', elements: [] },
		] as unknown as Parameters<UseThemeHandlersInput['setSlides']>[0];
		const input = {
			handlerRef: { current: handler as PptxHandler },
			serializeSlides: vi
				.fn<() => Promise<Uint8Array | null>>()
				.mockResolvedValue(new Uint8Array()),
			setContent: vi.fn<() => void>(),
			onContentChange: undefined,
			setTheme: vi.fn<() => void>(),
			setSlideMasters: vi.fn<() => void>(),
			slideMasters: [],
			history: { markDirty: vi.fn<() => void>() } as unknown as UseThemeHandlersInput['history'],
			setSlides: vi.fn((updater: unknown) =>
				typeof updater === 'function' ? updater(slides) : updater,
			),
			templateElementsBySlideId: {},
			setTemplateElementsBySlideId: vi.fn<() => void>(),
			theme: { name: 'Office', fontScheme: FONTS } as UseThemeHandlersInput['theme'],
			bumpHistory: vi.fn<() => void>(),
		} satisfies UseThemeHandlersInput;
		return input;
	}

	it('handleApplyThemeData re-themes the live slides without serialising or reloading', async () => {
		const applyTheme = vi.fn().mockResolvedValue(undefined);
		const input = inPlaceInput({ applyTheme });
		const { handleApplyThemeData } = useThemeHandlers(input);

		await handleApplyThemeData({ accent1: '#90C226' } as never, FONTS!, 'Facet');

		expect(applyTheme).toHaveBeenCalledWith({ accent1: '#90C226' }, FONTS, 'Facet');
		expect(input.serializeSlides).not.toHaveBeenCalled();
		expect(input.setContent).not.toHaveBeenCalled();
		expect(input.setSlides).toHaveBeenCalledWith(expect.any(Function));
		expect(input.bumpHistory).toHaveBeenCalledOnce();
		expect(input.history.markDirty).toHaveBeenCalledWith();
	});

	it('handleUpdateThemeFontScheme re-resolves fonts in place as one undo step', async () => {
		const updateThemeFontScheme = vi.fn().mockResolvedValue(undefined);
		const input = inPlaceInput({ updateThemeFontScheme });
		const { handleUpdateThemeFontScheme } = useThemeHandlers(input);
		const next = {
			majorFont: { latin: 'Georgia' },
			minorFont: { latin: 'Verdana' },
		} as typeof FONTS;

		await handleUpdateThemeFontScheme(next!);

		expect(updateThemeFontScheme).toHaveBeenCalledWith(next);
		expect(input.serializeSlides).not.toHaveBeenCalled();
		expect(input.setContent).not.toHaveBeenCalled();
		expect(input.setSlides).toHaveBeenCalledWith(expect.any(Function));
		expect(input.setTemplateElementsBySlideId).toHaveBeenCalledWith(expect.any(Function));
		expect(input.bumpHistory).toHaveBeenCalledOnce();
	});

	it('a fonts pick writes only the font scheme, leaving the colours and the name as authored', async () => {
		// `applyTheme` would also rewrite the colour scheme (a `sysClr` slot becomes `srgbClr`) and
		// copy the primary scheme into every theme part.
		const text = {
			type: 'text',
			id: 'body',
			x: 0,
			y: 0,
			width: 20,
			height: 20,
			textSegments: [
				{ text: 'Body', style: { fontFamily: 'Calibri', latinFontThemeToken: '+mn-lt' } },
			],
		} as unknown as PptxElement;
		let slides = [{ id: 'slide', slideNumber: 1, elements: [text] }] as PptxSlide[];
		let templates: Record<string, PptxElement[]> = { slide: [text] };
		let theme: Record<string, unknown> | null = {
			name: 'Kept',
			colorScheme: PRESET_THEMES[0].colorScheme,
		};
		const handler = {
			applyTheme: vi.fn(),
			updateThemeColorScheme: vi.fn(),
			updateThemeName: vi.fn(),
			updateThemeFontScheme: vi.fn().mockResolvedValue(undefined),
		};
		const input: UseThemeHandlersInput = {
			...inPlaceInput(handler as unknown as Partial<PptxHandler>),
			theme: { fontScheme: { minorFont: { latin: 'Calibri' } } },
			setSlides: vi.fn((update) => {
				slides = typeof update === 'function' ? update(slides) : update;
			}),
			templateElementsBySlideId: templates,
			setTemplateElementsBySlideId: vi.fn((update) => {
				templates = typeof update === 'function' ? update(templates) : update;
			}),
			setTheme: vi.fn((update) => {
				theme = typeof update === 'function' ? update(theme) : update;
			}),
		};
		const fontScheme = { majorFont: { latin: 'Arial' }, minorFont: { latin: 'Arial' } };

		await useThemeHandlers(input).handleUpdateThemeFontScheme(fontScheme);

		expect(handler.updateThemeFontScheme).toHaveBeenCalledExactlyOnceWith(fontScheme);
		expect(handler.updateThemeColorScheme).not.toHaveBeenCalled();
		expect(handler.updateThemeName).not.toHaveBeenCalled();
		expect(handler.applyTheme).not.toHaveBeenCalled();
		expect(input.bumpHistory).toHaveBeenCalledOnce();
		expect(input.history.markDirty).toHaveBeenCalledOnce();
		const face = (element: PptxElement | undefined) =>
			(element as { textSegments: Array<{ style: { fontFamily: string } }> }).textSegments[0].style
				.fontFamily;
		expect(face(slides[0].elements[0])).toBe('Arial');
		expect(face(templates.slide[0])).toBe('Arial');
		expect(theme).toStrictEqual({
			name: 'Kept',
			colorScheme: PRESET_THEMES[0].colorScheme,
			fontScheme,
		});
		expect(input.serializeSlides).not.toHaveBeenCalled();
		expect(input.setContent).not.toHaveBeenCalled();
	});
});
