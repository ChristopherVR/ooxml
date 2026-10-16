<script lang="ts">
	import { buildRibbonGallery } from 'ooxml-ui/pptx';
	import type {
		RibbonGalleryPickEvent,
		RibbonGalleryPlacement,
		RibbonGalleryPreviewEvent,
	} from 'ooxml-ui/pptx';
	import { useTranslator } from '../../../../i18n/context';
	import { refocusViewerRoot } from '../anchored-popup';
	import { strictTranslator } from './gallery-labels';
	import { useRibbonGalleryHost } from './ribbon-gallery-host';

	const { placement, chevronOnly = false, tagControl = true, oncommand }: { placement: RibbonGalleryPlacement; chevronOnly?: boolean; tagControl?: boolean; oncommand?: (command: 'formatBackground') => void } = $props();
	const host = useRibbonGalleryHost();
	const t = useTranslator();
	const translateLabel = strictTranslator(t);
	const descriptor = $derived(host ? host.build(placement.gallery) : buildRibbonGallery(placement.gallery, { element: null }));
	async function pick(event: Event): Promise<void> {
		refocusViewerRoot(event.currentTarget as HTMLElement);
		const { itemId, action } = (event as RibbonGalleryPickEvent).detail;
		const result = await host?.apply(placement.gallery, itemId, action);
		if (result?.kind === 'command') {
			oncommand?.(result.command);
		}
	}
	function preview(event: Event): void {
		host?.preview(placement.gallery, (event as RibbonGalleryPreviewEvent).detail.itemId);
	}
</script>
<pptx-ui-ribbon-gallery data-ribbon-control={tagControl ? placement.control : undefined} mode={placement.mode} chevron-only={chevronOnly || undefined}
	{descriptor} {translateLabel} disabled={!host} ongallery-pick={pick} ongallery-preview={preview} ongallery-preview-end={() => host?.endPreview()}></pptx-ui-ribbon-gallery>
