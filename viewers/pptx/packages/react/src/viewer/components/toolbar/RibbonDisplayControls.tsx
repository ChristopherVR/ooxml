import { RIBBON_DISPLAY_OPTIONS_LABEL_KEY, RIBBON_DISPLAY_REVEAL_LABEL_KEY } from 'ooxml-ui/pptx';
import type { RibbonDisplayController, RibbonDisplayState } from 'ooxml-ui/pptx';
import React from 'react';

import { useTranslation } from '../../i18n/use-translation';
import { cn } from '../../utils';

/**
 * The whole ribbon laid over the top of the viewer in full-screen mode. The page background is
 * solid underneath, so the slide does not show through; the card tint sits on top of it.
 */
export const RIBBON_PEEK_ROOT_CLASS =
	'absolute inset-x-0 top-0 z-[35] border-b border-border bg-background shadow-[0_8px_20px_rgb(0_0_0/0.2)] before:absolute before:inset-0 before:-z-10 before:bg-card before:content-[""]';

/**
 * The commands of "Show tabs only" laid over the slide, directly below the tab row. Opaque in the
 * same way as {@link RIBBON_PEEK_ROOT_CLASS}.
 */
export const RIBBON_OVERLAY_CLASS =
	'absolute inset-x-0 top-full z-[35] border-b border-border bg-background shadow-[0_8px_20px_rgb(0_0_0/0.2)] before:absolute before:inset-0 before:-z-10 before:bg-card before:content-[""]';

/** The slim strip that brings the ribbon back in full-screen mode. It keeps its place when peeked. */
export function RibbonRevealStrip({
	ribbon,
}: {
	ribbon: RibbonDisplayController;
}): React.ReactElement {
	const { t } = useTranslation();
	const label = t(RIBBON_DISPLAY_REVEAL_LABEL_KEY);
	return (
		<button
			type='button'
			data-pptx-chrome='ribbon-reveal'
			aria-label={label}
			title={label}
			onClick={() => ribbon.reveal()}
			className='flex h-[14px] w-full shrink-0 items-center justify-center border-b border-border bg-card text-[14px] leading-none tracking-[2px] text-muted-foreground hover:bg-accent hover:text-foreground'
		>
			···
		</button>
	);
}

/** The Ribbon Display Options button at the right end of the tab row. */
export function RibbonDisplayButton({
	ribbon,
	state,
}: {
	ribbon: RibbonDisplayController;
	state: RibbonDisplayState;
}): React.ReactElement {
	const { t } = useTranslation();
	const label = t(RIBBON_DISPLAY_OPTIONS_LABEL_KEY);
	return (
		<button
			type='button'
			aria-haspopup='menu'
			aria-expanded={state.menuOpen}
			aria-label={label}
			title={label}
			data-pptx-chrome='ribbon-display-options'
			onClick={() => ribbon.toggleMenu()}
			className={cn(
				'mx-1 inline-flex h-[22px] w-[22px] shrink-0 items-center justify-center rounded-sm text-muted-foreground hover:bg-accent hover:text-foreground',
				state.menuOpen && 'bg-accent text-foreground',
			)}
		>
			<svg
				viewBox='0 0 16 16'
				width='12'
				height='12'
				fill='none'
				stroke='currentColor'
				strokeWidth={1.5}
				strokeLinecap='round'
				strokeLinejoin='round'
				aria-hidden='true'
				className={cn(state.effectiveMode === 'tabs' && 'rotate-180')}
			>
				<path d='m4 6 4 4 4-4' />
			</svg>
		</button>
	);
}

/**
 * The Ribbon Display Options menu, right-aligned under the tab row. It sits in a zero-height
 * anchor placed after the tab row, so the horizontally scrolling tab row cannot clip it.
 */
export function RibbonDisplayMenu({
	ribbon,
	state,
}: {
	ribbon: RibbonDisplayController;
	state: RibbonDisplayState;
}): React.ReactElement {
	const { t } = useTranslation();
	return (
		<div className='relative h-0'>
			<div
				role='menu'
				aria-label={t(RIBBON_DISPLAY_OPTIONS_LABEL_KEY)}
				data-pptx-chrome='ribbon-display-menu'
				className='absolute right-1.5 top-0.5 z-50 flex min-w-[200px] flex-col gap-0.5 rounded-md border border-border bg-popover p-1 text-[12px] shadow-lg'
			>
				{state.menu.map((item) => (
					<button
						key={item.mode}
						type='button'
						role='menuitemradio'
						aria-checked={item.checked}
						data-ribbon-display={item.mode}
						onClick={() => ribbon.choose(item.mode)}
						className='flex w-full items-center gap-2 whitespace-nowrap rounded-sm px-2 py-1.5 text-left text-foreground hover:bg-accent/30'
					>
						<svg
							viewBox='0 0 16 16'
							width='12'
							height='12'
							fill='none'
							stroke='currentColor'
							strokeWidth={1.5}
							strokeLinecap='round'
							strokeLinejoin='round'
							aria-hidden='true'
							className={cn('shrink-0', !item.checked && 'invisible')}
						>
							<path d='m3 8.5 3 3 7-7' />
						</svg>
						{t(item.labelKey)}
					</button>
				))}
			</div>
		</div>
	);
}
