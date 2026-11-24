import {
	RIBBON_AUTO_COLLAPSE_MEDIA,
	createRibbonDisplayController,
	isRibbonCommandChange,
	isRibbonCommandClick,
} from 'ooxml-ui/pptx';
import type { RibbonDisplayController, RibbonDisplayMode } from 'ooxml-ui/pptx';
import type { RefObject } from 'react';
import { useEffect, useRef, useState } from 'react';

/** The ribbon and the reveal strip: a press inside them is not "outside the ribbon". */
const RIBBON_SURFACE = '[data-pptx-chrome="ribbon"], [data-pptx-chrome="ribbon-reveal"]';
/** The ribbon itself: commands executed inside it put a peeked ribbon away. */
const RIBBON_ROOT = '[data-pptx-chrome="ribbon"]';
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
	/** Options > Quick Access Toolbar > "Show Quick Access Toolbar". */
	quickAccessVisible?: boolean;
	/** Whether the host keeps the `quickAccessToolbar` panel available. */
	quickAccessAvailable?: boolean;
	/** Called when the user toggles Quick Access Toolbar from the Ribbon Display Options menu. */
	onQuickAccessVisibleChange?: ((visible: boolean) => void) | undefined;
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
	const {
		ribbonDisplay,
		collapseRibbonAutomatically,
		quickAccessVisible,
		quickAccessAvailable,
		rootRef,
	} = input;
	const onChangeRef = useRef(input.onRibbonDisplayChange);
	onChangeRef.current = input.onRibbonDisplayChange;
	const onQuickAccessRef = useRef(input.onQuickAccessVisibleChange);
	onQuickAccessRef.current = input.onQuickAccessVisibleChange;

	const [controller] = useState(() =>
		createRibbonDisplayController({
			initial: ribbonDisplay,
			onChange: (mode) => onChangeRef.current?.(mode),
			quickAccessVisible,
			quickAccessAvailable,
			onQuickAccessVisibleChange: (visible) => onQuickAccessRef.current?.(visible),
		}),
	);

	useEffect(() => {
		controller.setMode(ribbonDisplay);
	}, [controller, ribbonDisplay]);

	useEffect(() => {
		controller.setAutoCollapse(collapseRibbonAutomatically);
	}, [controller, collapseRibbonAutomatically]);

	useEffect(() => {
		if (quickAccessVisible !== undefined) {
			controller.setQuickAccessVisible(quickAccessVisible);
		}
	}, [controller, quickAccessVisible]);

	useEffect(() => {
		if (quickAccessAvailable !== undefined) {
			controller.setQuickAccessAvailable(quickAccessAvailable);
		}
	}, [controller, quickAccessAvailable]);

	useEffect(() => {
		const query = globalThis.matchMedia?.(RIBBON_AUTO_COLLAPSE_MEDIA);
		const sync = (): void => controller.setShortViewport(query?.matches ?? false);
		sync();
		query?.addEventListener('change', sync);
		return () => query?.removeEventListener('change', sync);
	}, [controller]);

	useEffect(() => {
		const pendingCommandClicks = new WeakSet<Event>();
		const inRibbon = (event: Event): boolean => {
			const path = typeof event.composedPath === 'function' ? event.composedPath() : [];
			if (path.some((node) => node instanceof Element && node.matches(RIBBON_ROOT))) {
				return true;
			}
			return event.target instanceof Element && event.target.closest(RIBBON_ROOT) !== null;
		};
		const onPointerDown = (event: PointerEvent): void => {
			const target = event.target instanceof Element ? event.target : null;
			if (!target?.closest(RIBBON_SURFACE)) {
				controller.dismiss();
			} else if (!target.closest(RIBBON_MENU_SURFACE)) {
				controller.setMenuOpen(false);
			}
		};
		const onClickCapture = (event: MouseEvent): void => {
			if (controller.getSnapshot().peek && inRibbon(event) && isRibbonCommandClick(event)) {
				pendingCommandClicks.add(event);
			}
		};
		const onClickBubble = (event: MouseEvent): void => {
			if (pendingCommandClicks.has(event)) {
				pendingCommandClicks.delete(event);
				controller.commandExecuted();
			}
		};
		const onChange = (event: Event): void => {
			if (controller.getSnapshot().peek && inRibbon(event) && isRibbonCommandChange(event)) {
				controller.commandExecuted();
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
		document.addEventListener('click', onClickCapture, true);
		document.addEventListener('click', onClickBubble);
		document.addEventListener('change', onChange);
		document.addEventListener('keydown', onKeyDown);
		return () => {
			document.removeEventListener('pointerdown', onPointerDown);
			document.removeEventListener('click', onClickCapture, true);
			document.removeEventListener('click', onClickBubble);
			document.removeEventListener('change', onChange);
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
