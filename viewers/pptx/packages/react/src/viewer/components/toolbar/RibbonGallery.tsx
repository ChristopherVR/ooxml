import { applyRibbonGalleryItem, buildRibbonGallery } from 'ooxml-ui/pptx';
import type {
	PptxUiRibbonGalleryElement,
	RibbonGalleryPickEvent,
	RibbonGalleryPlacement,
	RibbonGalleryPreviewEvent,
} from 'ooxml-ui/pptx';
import React, { useEffect, useMemo, useRef } from 'react';
import { useTranslation } from 'react-i18next';

import { useRibbonGalleryCommands } from '../ribbon-gallery-context';

export interface RibbonGalleryProps {
	/** Runs a command a gallery lists under its tiles ("Format Background..."). */
	onCommand?: (command: 'formatBackground') => void;
	placement: RibbonGalleryPlacement;
	icon?: React.ReactNode;
	chevronOnly?: boolean;
	tagControl?: boolean;
}
/** Native mutations/history stay in the host; shared owns the entire gallery view. */
export function RibbonGallery({
	placement,
	chevronOnly = false,
	tagControl = true,
	onCommand,
}: RibbonGalleryProps): React.ReactElement {
	const { t } = useTranslation();
	const ref = useRef<PptxUiRibbonGalleryElement>(null);
	const commands = useRibbonGalleryCommands();
	const context = commands?.context;
	const descriptor = useMemo(
		() => buildRibbonGallery(placement.gallery, context ?? { element: null }),
		[placement.gallery, context],
	);
	useEffect(() => {
		const host = ref.current;
		if (host) {
			host.descriptor = descriptor;
			host.translateLabel = t;
			host.disabled = !commands?.editable;
		}
	}, [descriptor, t, commands?.editable]);
	useEffect(() => {
		const host = ref.current;
		if (!host || !commands) {
			return;
		}
		const pick = (event: Event) => {
			if (!commands.editable) {
				return;
			}
			const { itemId, action } = (event as RibbonGalleryPickEvent).detail;
			const result = applyRibbonGalleryItem(placement.gallery, itemId, commands.context, action);
			if (result?.kind === 'command') {
				commands.endPreview?.();
				onCommand?.(result.command);
			} else if (result) {
				commands.dispatch(result);
			}
		};
		const preview = (event: Event) => {
			const { itemId } = (event as RibbonGalleryPreviewEvent).detail;
			commands.preview?.(applyRibbonGalleryItem(placement.gallery, itemId, commands.context));
		};
		const endPreview = () => commands.endPreview?.();
		host.addEventListener('gallery-pick', pick);
		host.addEventListener('gallery-preview', preview);
		host.addEventListener('gallery-preview-end', endPreview);
		return () => {
			host.removeEventListener('gallery-pick', pick);
			host.removeEventListener('gallery-preview', preview);
			host.removeEventListener('gallery-preview-end', endPreview);
		};
	}, [commands, placement.gallery, onCommand]);
	return (
		<pptx-ui-ribbon-gallery
			ref={ref}
			label={descriptor.label}
			mode={placement.mode}
			chevron-only={chevronOnly || undefined}
			data-ribbon-control={tagControl ? placement.control : undefined}
		/>
	);
}
