/**
 * RehearseTimingsHud: Timing overlay shown during rehearsal mode.
 * Displays per-slide time, total elapsed time, and a pause button.
 */
import React from 'react';
import { LuPause, LuPlay } from 'react-icons/lu';

import { useTranslation } from '../i18n/use-translation';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function formatMs(ms: number): string {
	const totalSeconds = Math.max(0, Math.floor(ms / 1000));
	const minutes = Math.floor(totalSeconds / 60);
	const seconds = totalSeconds % 60;
	return `${String(minutes)}:${String(seconds).padStart(2, '0')}`;
}

// ---------------------------------------------------------------------------
// Props
// ---------------------------------------------------------------------------

export interface RehearseTimingsHudProps {
	/** Time on the current slide (the shared rehearsal controller ticks it). */
	slideElapsedMs: number;
	/** Time since the rehearsal started, paused time excluded. */
	totalElapsedMs: number;
	/** Whether the rehearsal timer is paused. */
	paused: boolean;
	/** Toggle the paused state. */
	onTogglePause: () => void;
}

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

export function RehearseTimingsHud({
	slideElapsedMs: slideElapsed,
	totalElapsedMs: totalElapsed,
	paused,
	onTogglePause,
}: RehearseTimingsHudProps): React.ReactElement {
	const { t } = useTranslation();

	return (
		<div className='fixed bottom-4 left-4 z-[9999] flex items-center gap-3 rounded-lg bg-black/80 px-4 py-2 text-white shadow-xl backdrop-blur-sm select-none'>
			<div className='flex flex-col text-xs leading-tight'>
				<span className='text-muted-foreground'>{t('pptx.rehearse.slideTime')}</span>
				<span className='text-lg font-mono tabular-nums'>{formatMs(slideElapsed)}</span>
			</div>
			<div className='w-px h-8 bg-border' />
			<div className='flex flex-col text-xs leading-tight'>
				<span className='text-muted-foreground'>{t('pptx.rehearse.totalTime')}</span>
				<span className='text-lg font-mono tabular-nums'>{formatMs(totalElapsed)}</span>
			</div>
			<button
				type='button'
				onClick={onTogglePause}
				className='ml-1 rounded p-1.5 hover:bg-white/20 transition-colors'
				title={paused ? t('pptx.rehearse.resume') : t('pptx.rehearse.pause')}
				aria-label={paused ? t('pptx.rehearse.resume') : t('pptx.rehearse.pause')}
			>
				{paused ? <LuPlay className='w-4 h-4' /> : <LuPause className='w-4 h-4' />}
			</button>
		</div>
	);
}
