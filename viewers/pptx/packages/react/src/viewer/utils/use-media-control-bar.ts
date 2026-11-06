import { attachMediaControlBar, type MediaControlBar } from 'ooxml-ui/pptx';
import { useEffect, useRef } from 'react';

import { useTranslation } from '../i18n/use-translation';

/**
 * PowerPoint's Normal-view control bar for a video or audio shape: none at
 * rest, one below the shape while it is selected or hovered. The bar itself is
 * the shared `attachMediaControlBar`; React only mounts it on the element box
 * that owns the media and keeps `selected` current.
 */
export function useMediaControlBar(
	mediaRef: React.RefObject<HTMLVideoElement | HTMLAudioElement | null>,
	enabled: boolean,
	selected: boolean,
): void {
	const { t } = useTranslation();
	const barRef = useRef<MediaControlBar | null>(null);
	const selectedRef = useRef(selected);
	selectedRef.current = selected;

	useEffect(() => {
		const media = mediaRef.current;
		if (!enabled || !media) {
			return undefined;
		}
		const host =
			media.closest<HTMLElement>('[data-pptx-element]') ?? media.parentElement ?? undefined;
		if (!host) {
			return undefined;
		}
		const bar = attachMediaControlBar(host, media, { t, selected: selectedRef.current });
		barRef.current = bar;
		return () => {
			bar.dispose();
			barRef.current = null;
		};
	}, [mediaRef, enabled, t]);

	useEffect(() => {
		barRef.current?.setSelected(selected);
	}, [selected]);
}
