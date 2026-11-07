import type { DeckWebfontProgress } from 'ooxml-ui/pptx';
import React from 'react';

import { useTranslation } from '../i18n/use-translation';
import { cn } from '../utils';

interface LoadingStateProps {
	className?: string;
	/** Set while the load waits for the deck's webfonts (after a short delay). */
	fontProgress?: DeckWebfontProgress | null;
}

export function LoadingState({ className, fontProgress }: LoadingStateProps): React.ReactElement {
	const { t } = useTranslation();
	return (
		<div
			role='status'
			aria-live='polite'
			aria-atomic='true'
			className={cn(
				'h-full w-full flex items-center justify-center text-muted-foreground',
				className,
			)}
		>
			<div className='flex items-center gap-2'>
				<div
					className='animate-spin rounded-full h-6 w-6 border-b-2 border-primary'
					aria-hidden='true'
				/>
				<span>
					{fontProgress
						? t('pptx.viewer.loadingFonts', {
								loaded: fontProgress.loaded,
								total: fontProgress.total,
							})
						: t('pptx.viewer.loading')}
				</span>
			</div>
		</div>
	);
}
