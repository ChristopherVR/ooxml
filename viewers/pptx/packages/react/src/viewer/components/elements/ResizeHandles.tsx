import { getResizeHandleHitAreaStyle } from 'ooxml-ui/pptx';
import React from 'react';

import { useTranslation } from '../../i18n/use-translation';
import type { ShapeAdjustmentHandleDescriptor } from '../../types';
import { cn } from '../../utils';
import { CORNER_HANDLES, EDGE_HANDLES } from './resize-handle-classes';
import { startRotateKnobGesture } from './rotate-knob-gesture';
import { RotateHandleArtwork } from './RotateHandleArtwork';
import {
	RESIZE_ARTWORK,
	RESIZE_DEFAULT_CLASSES,
	ROTATE_ARTWORK,
	ROTATE_DEFAULT_CLASSES,
} from './selection-control-artwork';
import { useRotateHandlePlacement } from './use-rotate-handle-placement';

export { CORNER_HANDLES, EDGE_HANDLES } from './resize-handle-classes';

export interface ResizeHandlesProps {
	elementId: string;
	/**
	 * Every `a:avLst` handle the shape offers, not just the first: PowerPoint
	 * shows one amber diamond per adjustable parameter and presets routinely
	 * have several (`quadArrow` three, `callout3` four).
	 */
	adjustmentHandles: ShapeAdjustmentHandleDescriptor[];
	onResizePointerDown: (elementId: string, e: React.MouseEvent, handle: string) => void;
	onAdjustmentPointerDown: (
		elementId: string,
		e: React.MouseEvent,
		descriptor: ShapeAdjustmentHandleDescriptor,
	) => void;
	/** Whether to force pointerEvents: "auto" on buttons (needed inside pointer-events:none containers). */
	forcePointerEvents?: boolean;
	/** Current element rotation in degrees (the rotate-handle drag baseline). */
	rotation?: number;
	/** Element transform sans rotation (flips/skews); the live-preview base. */
	nonRotationTransform?: string;
	/** Commit a new rotation (degrees) for the element. Omit to hide the handle. */
	onRotate?: (elementId: string, rotationDeg: number) => void;
}

/**
 * Touch-action: none stops the browser from claiming touch gestures (scroll /
 * pinch-zoom) over a handle so a finger drag becomes a resize. Applied to every
 * handle button alongside the pointer-down wiring below.
 */
const HANDLE_TOUCH_ACTION = {
	touchAction: 'none' as const,
	// Keep both the visible indicator and its expanded hit area in screen pixels.
	// Only the button scales; its center and the overlay remain in slide coordinates.
	scale: 'var(--pptx-handle-inverse-scale, 1)',
};

export function ResizeHandles({
	elementId,
	adjustmentHandles,
	onResizePointerDown,
	onAdjustmentPointerDown,
	forcePointerEvents,
	rotation,
	nonRotationTransform,
	onRotate,
}: ResizeHandlesProps) {
	const { t } = useTranslation();
	const rotateRef = useRotateHandlePlacement(elementId, Boolean(onRotate));
	const peStyle = forcePointerEvents
		? { ...HANDLE_TOUCH_ACTION, pointerEvents: 'auto' as const }
		: HANDLE_TOUCH_ACTION;
	// The bounded child owns the press, which still bubbles through its semantic
	// button for focus, mouse/touch dispatch and pointer capture. Neither the
	// indicator nor an unbounded button box may steal a neighboring handle.
	const resizeStyle = { ...HANDLE_TOUCH_ACTION, pointerEvents: 'none' as const };

	// Touch/pen presses start the resize via Pointer Events (mouse keeps using
	// onMouseDown so desktop behaviour is unchanged and never double-fires). The
	// pointer is captured so the gesture keeps tracking even if the finger drifts
	// off the small handle.
	const handleResizePointer = (e: React.PointerEvent, handle: string) => {
		if (e.pointerType === 'mouse') {
			return;
		}
		e.stopPropagation();
		(e.currentTarget as Element).setPointerCapture?.(e.pointerId);
		onResizePointerDown(elementId, e, handle);
	};

	// ── Rotate handle (self-contained) ───────────────────────────────────
	// Shared gesture, live preview on the element's transform, commit on release.
	const startRotate = (button: HTMLElement, pointer: React.PointerEvent): void =>
		startRotateKnobGesture({
			button,
			elementId,
			pointer,
			pointerId: pointer.pointerId,
			rotation: rotation ?? 0,
			...(nonRotationTransform ? { nonRotationTransform } : {}),
			...(onRotate ? { onRotate } : {}),
		});

	return (
		<>
			{/* The semantic frame keeps the bounded hit region; artwork never owns presses. */}
			{[...CORNER_HANDLES, ...EDGE_HANDLES].map(({ handle, posClass, cursor }) => {
				const corner = handle.length === 2;
				const styles =
					RESIZE_ARTWORK[
						corner ? 'corner' : handle === 'n' || handle === 's' ? 'horizontal' : 'vertical'
					];
				return (
					<button
						data-export-ignore='true'
						key={handle}
						type='button'
						aria-label={t('pptx.selectionOverlay.resize', { handle })}
						data-pptx-handle-kind='resize'
						data-pptx-compact
						className={cn(
							'absolute z-10 group border-0 bg-transparent p-0',
							RESIZE_DEFAULT_CLASSES,
							posClass,
							cursor,
						)}
						style={{ ...resizeStyle, ...styles.frame }}
						onPointerDown={(e) => handleResizePointer(e, handle)}
						onMouseDown={(e) => {
							e.stopPropagation();
							onResizePointerDown(elementId, e, handle);
						}}
					>
						<div
							data-pptx-handle-artwork
							aria-hidden='true'
							className='border shadow'
							style={styles.artwork}
						/>
						{/* Invisible expanded hit area. The larger touch-friendly pad is
						    `max-md:`-only (this repo's touch proxy, see toolbar-constants.tsx):
						    unconditionally applying it also on a fine-pointer desktop let an
						    edge handle's 6-8px inward reach cover nearby in-shape content
						    (e.g. a chart title sitting close to the top edge), intercepting
						    clicks meant for that content instead of the handle. */}
						<div
							data-pptx-handle-hit
							className={cn(
								'absolute -inset-px pointer-events-auto [--pptx-handle-hit-inset:-1px]',
								corner
									? 'max-md:-inset-1.5 max-md:[--pptx-handle-hit-inset:-6px]'
									: 'max-md:-inset-2 max-md:[--pptx-handle-hit-inset:-8px]',
							)}
							style={getResizeHandleHitAreaStyle(handle)}
						/>
					</button>
				);
			})}

			{/* Rotate is separate from the North resize target. The placement hook
			    keeps its complete pointer target inside the visible canvas. */}
			{onRotate ? (
				<button
					data-export-ignore='true'
					ref={rotateRef}
					type='button'
					aria-label={t('pptx.selectionOverlay.rotate')}
					data-pptx-handle-kind='rotate'
					data-pptx-compact
					className={cn(
						'absolute left-1/2 top-0 -translate-x-1/2 -translate-y-1/2 z-20 border-0 bg-transparent p-0 cursor-grab active:cursor-grabbing',
						ROTATE_DEFAULT_CLASSES,
					)}
					style={{
						...peStyle,
						...ROTATE_ARTWORK.frame,
						top: 'calc(-24px * var(--pptx-handle-inverse-scale, 1))',
					}}
					onPointerDown={(e) => {
						// Keep the same subpixel coordinates as pointermove; legacy
						// mousedown rounds them and can shift a short shape's anchor.
						e.stopPropagation();
						startRotate(e.currentTarget, e);
					}}
					onMouseDown={(e) => {
						e.stopPropagation();
					}}
				>
					<RotateHandleArtwork />
				</button>
			) : null}

			{/* Shape adjustment handles (yellow diamonds), one per `a:avLst` guide.
			    Every one carries the SAME accessible name: `playwright.config.ts`
			    lists `aria-label="Adjust shape"` as part of the framework-neutral
			    contract all five viewers emit. The offsets centre the 10px diamond
			    on the element-local point shared measured off the preset geometry. */}
			{adjustmentHandles.map((adjH) => (
				<button
					data-export-ignore='true'
					key={adjH.key}
					type='button'
					aria-label={t('pptx.canvas.adjustShape')}
					data-pptx-adjust-key={adjH.key}
					data-pptx-handle-kind='adjust'
					data-pptx-compact
					className='absolute h-2.5 w-2.5 max-md:h-4 max-md:w-4 rotate-45 border border-amber-700 bg-amber-300 shadow z-10'
					style={{
						left: adjH.left - 5,
						top: adjH.top - 5,
						cursor: adjH.cursor,
						...HANDLE_TOUCH_ACTION,
						...(forcePointerEvents ? { pointerEvents: 'auto' as const } : {}),
					}}
					onPointerDown={(e) => {
						if (e.pointerType === 'mouse') {
							return;
						}
						e.stopPropagation();
						(e.currentTarget as Element).setPointerCapture?.(e.pointerId);
						onAdjustmentPointerDown(elementId, e, adjH);
					}}
					onMouseDown={(e) => {
						e.stopPropagation();
						onAdjustmentPointerDown(elementId, e, adjH);
					}}
				/>
			))}
		</>
	);
}
