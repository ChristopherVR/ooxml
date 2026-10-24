/**
 * PrintDialog: Full-featured print dialog for PowerPoint presentations.
 *
 * Options: print what (slides/handouts/notes/outline), slides per page,
 * orientation, colour mode, frame slides, slide range.
 */
import {
	computeSlideCount,
	computeSlideIndices,
	effectiveOrientation as resolveOrientation,
	estimatePageCount,
	getDensePanelTouchTargetPx,
	resolveDefaultPrintSettings,
} from 'ooxml-ui/pptx';
import React, { useCallback, useMemo, useState } from 'react';
import { LuPrinter, LuX } from 'react-icons/lu';

import { useModalDismissDrag } from '../hooks';
import { useIsMobile } from '../hooks/useIsMobile';
import { useTranslation } from '../i18n/use-translation';
import { DialogFooter } from './DialogFooter';
import { PrintPreview, NotesPagePreview } from './print';
import type {
	PrintWhat,
	PrintOrientation,
	PrintColorMode,
	HandoutSlidesPerPage,
	PrintSlideRange,
	PrintDialogProps,
} from './print-dialog-types';
import { HANDOUT_OPTIONS } from './print-dialog-types';
import { PrintSettingsPanel } from './PrintSettingsPanel';
import { useViewerOptionsContext } from './viewer-options-context';
import { ViewerRegion, useRegionRenderer } from './ViewerRegion';

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

export function PrintDialog({
	open,
	onClose,
	onPrint,
	slides,
	activeSlideIndex,
	defaultSlidesPerPage,
	defaultFrameSlides,
}: PrintDialogProps): React.ReactElement | null {
	const { t } = useTranslation();
	const regionRenderer = useRegionRenderer('dialog.print');
	const { panelStyle, handlers: dragHandlers } = useModalDismissDrag(onClose);
	const { viewportWidth } = useIsMobile();
	const touchTargetPx = getDensePanelTouchTargetPx(viewportWidth);
	const closeBtnStyle = { minWidth: touchTargetPx, minHeight: touchTargetPx };

	// File > Options > Advanced > Print. `undefined` (Options > "Use the most
	// recently used print settings") keeps this dialog's own sticky in-session
	// state, since its component instance stays mounted (returns `null` while
	// `open` is false) rather than remounting, so React state already IS the
	// "most recent settings" store across opens within one viewer session.
	const printOptionDefaults = resolveDefaultPrintSettings(useViewerOptionsContext());

	// ── State ───────────────────────────────────────────────────────────
	const [printWhat, setPrintWhat] = useState<PrintWhat>(printOptionDefaults?.printWhat ?? 'slides');
	const [orientation, setOrientation] = useState<PrintOrientation>('landscape');
	const [colorMode, setColorMode] = useState<PrintColorMode>(
		printOptionDefaults?.colorMode ?? 'color',
	);
	const [frameSlides, setFrameSlides] = useState(
		printOptionDefaults?.frameSlides ?? defaultFrameSlides ?? false,
	);
	const [slidesPerPage, setSlidesPerPage] = useState<HandoutSlidesPerPage>(
		(HANDOUT_OPTIONS.includes(defaultSlidesPerPage as HandoutSlidesPerPage)
			? defaultSlidesPerPage
			: 6) as HandoutSlidesPerPage,
	);
	const [slideRange, setSlideRange] = useState<PrintSlideRange>('all');
	const [customFrom, setCustomFrom] = useState(1);
	const [customTo, setCustomTo] = useState(slides.length);

	// ── Derived ─────────────────────────────────────────────────────────
	// The shared print helpers every binding's dialog uses (the same ones the
	// print job plan validates against).
	const effectiveOrientation = useMemo<PrintOrientation>(
		() => resolveOrientation(printWhat, orientation),
		[printWhat, orientation],
	);

	const slideCount = useMemo(
		() => computeSlideCount(slideRange, slides.length, customFrom, customTo),
		[slideRange, slides.length, customFrom, customTo],
	);

	const pageCount = useMemo(
		() => estimatePageCount(printWhat, slideCount, slidesPerPage),
		[printWhat, slideCount, slidesPerPage],
	);

	const previewSlideIndices = useMemo<number[]>(
		() => computeSlideIndices(slideRange, activeSlideIndex, slides.length, customFrom, customTo),
		[slideRange, activeSlideIndex, customFrom, customTo, slides.length],
	);

	// ── Handlers ────────────────────────────────────────────────────────
	const handlePrint = useCallback(() => {
		onPrint({
			printWhat,
			orientation: effectiveOrientation,
			colorMode,
			frameSlides,
			slidesPerPage,
			slideRange,
			customRangeFrom: Math.max(1, Math.min(customFrom, slides.length)),
			customRangeTo: Math.max(1, Math.min(customTo, slides.length)),
			// Options > Advanced > "Print scale to fit"; no dialog control of its
			// own (PowerPoint keeps this an Options default, not a per-job choice).
			scaleToFit: printOptionDefaults?.scaleToFit ?? true,
		});
	}, [
		printWhat,
		effectiveOrientation,
		colorMode,
		frameSlides,
		slidesPerPage,
		slideRange,
		customFrom,
		customTo,
		slides.length,
		printOptionDefaults?.scaleToFit,
		onPrint,
	]);

	if (!open) {
		return null;
	}

	// ── Render ──────────────────────────────────────────────────────────
	return (
		<div
			style={{ zIndex: 1200 }}
			className='fixed inset-0 flex items-center justify-center bg-black/60 backdrop-blur-sm'
		>
			<div
				style={panelStyle}
				role='dialog'
				aria-modal='true'
				aria-label={t('pptx.print.title')}
				className='w-[780px] max-h-[90vh] rounded-xl border border-border bg-background shadow-2xl flex flex-col max-md:fixed max-md:inset-x-0 max-md:bottom-0 max-md:top-auto max-md:w-full max-md:max-w-none max-md:max-h-[88dvh] max-md:rounded-t-2xl max-md:rounded-b-none max-md:border-x-0 max-md:border-b-0 max-md:pb-[max(env(safe-area-inset-bottom),0px)]'
			>
				{/* Header — also a swipe-down-to-dismiss grab region on touch. */}
				<div
					{...dragHandlers}
					className='flex items-center justify-between px-5 py-4 border-b border-border touch-none'
				>
					<div className='flex items-center gap-2'>
						<LuPrinter className='w-4 h-4 text-muted-foreground' />
						<h2 className='text-sm font-semibold text-foreground'>{t('pptx.print.title')}</h2>
					</div>
					<button
						type='button'
						onClick={onClose}
						style={closeBtnStyle}
						className='p-1 rounded hover:bg-accent text-muted-foreground hover:text-foreground transition-colors inline-flex items-center justify-center'
					>
						<LuX className='w-4 h-4' />
					</button>
				</div>

				{regionRenderer ? (
					// The frame (backdrop, title bar, close, dismiss) stays the viewer's; the host replaces the body.
					<div className='min-h-0 flex-1 overflow-y-auto'>
						<ViewerRegion
							region='dialog.print'
							renderer={regionRenderer}
							close={onClose}
							style={{ width: '100%', minHeight: '8rem' }}
						/>
					</div>
				) : (
					<>
						{/* Body */}
						<div className='flex-1 overflow-y-auto px-5 py-4 flex gap-5 max-md:flex-col'>
							{/* Left: Settings */}
							<PrintSettingsPanel
								printWhat={printWhat}
								onPrintWhatChange={setPrintWhat}
								orientation={orientation}
								onOrientationChange={setOrientation}
								colorMode={colorMode}
								onColorModeChange={setColorMode}
								frameSlides={frameSlides}
								onFrameSlidesChange={setFrameSlides}
								slidesPerPage={slidesPerPage}
								onSlidesPerPageChange={setSlidesPerPage}
								slideRange={slideRange}
								onSlideRangeChange={setSlideRange}
								customFrom={customFrom}
								onCustomFromChange={setCustomFrom}
								customTo={customTo}
								onCustomToChange={setCustomTo}
								totalSlides={slides.length}
								activeSlideIndex={activeSlideIndex}
							/>

							{/* Right: Preview */}
							{printWhat === 'handouts' && (
								<div className='w-[230px] shrink-0 border-l border-border pl-4 overflow-y-auto max-md:w-full max-md:border-l-0 max-md:border-t max-md:pl-0 max-md:pt-3'>
									<PrintPreview
										slideIndices={previewSlideIndices}
										slidesPerPage={slidesPerPage}
										orientation={effectiveOrientation}
										frameSlides={frameSlides}
									/>
								</div>
							)}
							{printWhat === 'notes' && (
								<div className='w-[230px] shrink-0 border-l border-border pl-4 overflow-y-auto max-md:w-full max-md:border-l-0 max-md:border-t max-md:pl-0 max-md:pt-3'>
									<NotesPagePreview
										slideIndices={previewSlideIndices}
										slides={slides}
										frameSlides={frameSlides}
									/>
								</div>
							)}
						</div>

						{/* Footer */}
						<div className='flex items-center justify-between px-5 py-3 border-t border-border'>
							<span className='text-xs text-muted-foreground'>
								{t('pptx.print.pageEstimate', {
									pages: pageCount,
									slides: slideCount,
								})}
							</span>
							<div className='flex gap-2'>
								<DialogFooter
									actions={[
										{ id: 'cancel', label: t('pptx.common.cancel') },
										{
											id: 'print',
											label: t('pptx.print.printButton'),
											variant: 'primary',
											icon: 'print',
										},
									]}
									onAction={(id) => (id === 'print' ? handlePrint() : onClose())}
								/>
							</div>
						</div>
					</>
				)}
			</div>
		</div>
	);
}
