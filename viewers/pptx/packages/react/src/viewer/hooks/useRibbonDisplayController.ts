import { RIBBON_AUTO_COLLAPSE_MEDIA, createRibbonDisplayController } from 'ooxml-ui/pptx';
import type { RibbonDisplayController, RibbonDisplayMode } from 'ooxml-ui/pptx';
import type { RefObject } from 'react';
import { useEffect, useRef, useState } from 'react';

/** The ribbon and the reveal strip: a press inside them is not "outside the ribbon". */
const RIBBON_SURFACE = '[data-pptx-chrome="ribbon"], [data-pptx-chrome="ribbon-reveal"]';
/** The menu button and its menu: a press here does not close the menu. */
const RIBBON_MENU_SURFACE =
	'[data-pptx-chrome="ribbon-display-options"], [data-pptx-chrome="ribbon-display-menu"]';

export interface RibbonDisplayControllerInput {
	/** The host's `ribbonDisplay` prop: the starting mode, and later changes from the host. */
	ribbonDisplay: RibbonDisplayMode | undefined;
	/** The host's `onRibbonDisplayChange`: fires only when the user picks a mode. */
	onRibbonDisplayChange: ((mode: RibbonDisplayMode) => void) | undefined;
	/** Options > General > "Collapse the ribbon automatically". */
	collapseRibbonAutomatically: boolean;
	/** The viewer's root element: key presses inside it reach the ribbon's shortcuts. */
	rootRef: RefObject<HTMLElement | null>;
}

/**
 * The viewer's one Ribbon Display Options controller. It is created once (never per render),
 * forwards user choices to the latest `onRibbonDisplayChange`, follows the host's
 * `ribbonDisplay` prop, the automatic-collapse option and the short-viewport media query, and
 * owns the document gestures (press outside, Ctrl+F1, Escape). Every listener is removed on
 * unmount.
 */
export function useRibbonDisplayController(
	input: RibbonDisplayControllerInput,
): RibbonDisplayController {
	const { ribbonDisplay, collapseRibbonAutomatically, rootRef } = input;
	const onChangeRef = useRef(input.onRibbonDisplayChange);
	onChangeRef.current = input.onRibbonDisplayChange;

	const [controller] = useState(() =>
		createRibbonDisplayController({
			initial: ribbonDisplay,
			onChange: (mode) => onChangeRef.current?.(mode),
		}),
	);

	useEffect(() => {
		controller.setMode(ribbonDisplay);
	}, [controller, ribbonDisplay]);

	useEffect(() => {
		controller.setAutoCollapse(collapseRibbonAutomatically);
	}, [controller, collapseRibbonAutomatically]);

	useEffect(() => {
		const query = globalThis.matchMedia?.(RIBBON_AUTO_COLLAPSE_MEDIA);
		const sync = (): void => controller.setShortViewport(query?.matches ?? false);
		sync();
		query?.addEventListener('change', sync);
		return () => query?.removeEventListener('change', sync);
	}, [controller]);

	useEffect(() => {
		const onPointerDown = (event: PointerEvent): void => {
			const target = event.target instanceof Element ? event.target : null;
			if (!target?.closest(RIBBON_SURFACE)) {
				controller.dismiss();
			} else if (!target.closest(RIBBON_MENU_SURFACE)) {
				controller.setMenuOpen(false);
			}
		};
		const onKeyDown = (event: KeyboardEvent): void => {
			const target = event.target;
			const root = rootRef.current;
			// Only this viewer's shortcuts: the focus is inside it, or nowhere in particular.
			if (target === document.body || (root && target instanceof Node && root.contains(target))) {
				controller.handleKeydown(event);
			}
		};
		document.addEventListener('pointerdown', onPointerDown);
		document.addEventListener('keydown', onKeyDown);
		return () => {
			document.removeEventListener('pointerdown', onPointerDown);
			document.removeEventListener('keydown', onKeyDown);
		};
	}, [controller, rootRef]);

	return controller;
}

/**
 * The controller a ribbon uses: the viewer's when one is passed, otherwise a private one that
 * stays in the default mode (a ribbon rendered on its own, as in unit tests).
 */
export function useRibbonDisplayFallback(
	controller: RibbonDisplayController | undefined,
): RibbonDisplayController {
	const [own] = useState(() => createRibbonDisplayController());
	return controller ?? own;
}
