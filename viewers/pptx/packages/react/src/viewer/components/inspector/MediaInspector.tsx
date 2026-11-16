import {
	MEDIA_TRIM_SECTIONS,
	mediaTrimResetPatch,
	mediaTrimResetState,
	trimmedMediaDurationMs,
	validateMediaTrimRange,
} from 'ooxml-ui/pptx';
import type { MediaTrimRangeError } from 'ooxml-ui/pptx';
import type { MediaPptxElement, PptxElement } from 'pptx-viewer-core';
import React, { useCallback, useMemo } from 'react';
import { LuClock, LuScissors } from 'react-icons/lu';

import { useTranslation } from '../../i18n/use-translation';
import { InspectorFields } from './InspectorFields';

// ---------------------------------------------------------------------------
// Style constants (matching InspectorPane)
// ---------------------------------------------------------------------------

const HEADING = 'text-[11px] uppercase tracking-wide text-muted-foreground';
const CARD = 'rounded border border-border bg-card p-2 space-y-2';
const BTN = 'rounded bg-muted hover:bg-accent px-2 py-1 text-[11px] transition-colors';
const LABEL_CLS = 'flex items-center justify-between gap-2';
const LABEL_TEXT = 'text-muted-foreground';

// ---------------------------------------------------------------------------
// Time conversion utilities (exported for testing)
// ---------------------------------------------------------------------------

/**
 * Convert milliseconds to a mm:ss display string.
 * E.g. 65000 → "01:05", 3661000 → "61:01" (minutes can exceed 59).
 */
export function msToMmSs(ms: number): string {
	if (!Number.isFinite(ms) || ms < 0) {
		return '00:00';
	}
	const totalSeconds = Math.round(ms / 1000);
	const minutes = Math.floor(totalSeconds / 60);
	const seconds = totalSeconds % 60;
	return `${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
}

const TRIM_ERROR_KEYS: Record<MediaTrimRangeError, string> = {
	negative: 'pptx.media.trimErrorNegative',
	startAfterEnd: 'pptx.media.trimErrorStartAfterEnd',
	beyondDuration: 'pptx.media.trimErrorBeyondDuration',
};

/**
 * Validate a trim range (`trimEndMs` is `p14:trim/@end`'s distance from the
 * clip's tail; the maths lives in shared). Returns an error i18n key or null.
 */
export function validateTrimRange(
	trimStartMs: number,
	trimEndMs: number,
	durationMs: number,
): string | null {
	const error = validateMediaTrimRange(trimStartMs, trimEndMs, durationMs);
	return error ? TRIM_ERROR_KEYS[error] : null;
}

// ---------------------------------------------------------------------------
// MediaTrimInspector: focused trim UI component
// ---------------------------------------------------------------------------

export interface MediaInspectorProps {
	element: MediaPptxElement;
	canEdit: boolean;
	durationSeconds: number;
	onUpdateElement: (updates: Partial<PptxElement>) => void;
}

export function MediaInspector({
	element,
	canEdit,
	durationSeconds,
	onUpdateElement,
}: MediaInspectorProps): React.ReactElement {
	const { t } = useTranslation();

	const trimStartMs = element.trimStartMs ?? 0;
	const trimEndMs = element.trimEndMs ?? 0;
	const durationMs = durationSeconds * 1000;

	// Auto-calculate trimmed duration
	const trimmedDuration = useMemo(
		(): string => msToMmSs(trimmedMediaDurationMs(trimStartMs, trimEndMs, durationMs)),
		[trimStartMs, trimEndMs, durationMs],
	);
	const validationError = useMemo(
		() => validateTrimRange(trimStartMs, trimEndMs, durationMs),
		[trimStartMs, trimEndMs, durationMs],
	);

	const resetTrimState = mediaTrimResetState({ trimStartMs, trimEndMs }, canEdit);

	const handleResetTrim = useCallback((): void => {
		onUpdateElement(mediaTrimResetPatch());
	}, [onUpdateElement]);

	return (
		<div className={CARD}>
			<div className={HEADING}>
				<LuScissors className='inline w-3 h-3 mr-1' />
				{t('pptx.media.trim')}
			</div>

			{/* Start and end times, bounded by the live duration: shared descriptors. */}
			<div className='grid grid-cols-2 gap-1.5 text-[11px]'>
				<InspectorFields
					sections={MEDIA_TRIM_SECTIONS}
					target={{ media: element, durationMs }}
					disabled={!canEdit}
					onPatch={onUpdateElement}
				/>
			</div>

			{/* Trimmed duration display */}
			<div className={LABEL_CLS}>
				<span className={`text-[11px] ${LABEL_TEXT}`}>
					<LuClock className='inline w-3 h-3 mr-1' />
					{t('pptx.media.trimmedDuration')}
				</span>
				<span className='text-[11px] tabular-nums font-medium'>{trimmedDuration}</span>
			</div>

			{/* Validation error */}
			{validationError && <div className='text-[10px] text-red-400'>{t(validationError)}</div>}

			{/* Reset trim button */}
			{resetTrimState.visible && (
				<button type='button' className={`${BTN} w-full text-center`} onClick={handleResetTrim}>
					{t('pptx.media.resetTrim')}
				</button>
			)}
		</div>
	);
}
