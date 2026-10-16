/**
 * The reactive wiring behind `RibbonGallery.vue`: the descriptor for one
 * gallery (rebuilt whenever the host context changes), its inline strip, the
 * translated labels, and the pick handler. Every decision is the shared
 * `buildRibbonGallery` / `applyRibbonGalleryItem`; this only adapts them to
 * Vue reactivity and the viewer's gallery host.
 */
import {
	applyRibbonGalleryItem,
	buildRibbonGallery,
	galleryHasItems,
	galleryItemLabel,
	inlineGalleryItems,
} from 'ooxml-ui/pptx';
import type {
	RibbonGalleryApplyResult,
	RibbonGalleryDescriptor,
	RibbonGalleryId,
	RibbonGalleryItem,
	RibbonGallerySection,
} from 'ooxml-ui/pptx';
import { computed } from 'vue';
import type { ComputedRef } from 'vue';
import { useI18n } from 'vue-i18n';

import { useRibbonGalleryHost } from '../../composables/useRibbonGalleryHost';

type Translate = (key: string, params?: Readonly<Record<string, string | number>>) => string;

export interface UseRibbonGalleryResult {
	descriptor: ComputedRef<RibbonGalleryDescriptor>;
	/** True when the trigger cannot open (selection does not take it, or no tiles). */
	disabled: ComputedRef<boolean>;
	inlineItems: ComputedRef<RibbonGalleryItem[]>;
	/** Translated gallery name. */
	title: ComputedRef<string>;
	itemLabel: (item: RibbonGalleryItem) => string;
	sectionTitle: (section: RibbonGallerySection) => string | undefined;
	/**
	 * Apply `itemId` (with the context menu `action` chosen on the tile, if any) and dispatch the
	 * result; returns the result when something was picked, so a caller can run a command.
	 */
	pick: (itemId: string, action?: string) => RibbonGalleryApplyResult | null;
	/** Hover preview of `itemId`; `endPreview` puts the slides back. */
	preview: (itemId: string) => void;
	endPreview: () => void;
	translate: Translate;
}

export function useRibbonGallery(gallery: () => RibbonGalleryId): UseRibbonGalleryResult {
	const { t } = useI18n();
	const host = useRibbonGalleryHost();
	const translate: Translate = (key, params) => (params ? t(key, { ...params }) : t(key));
	const orFallback = (key: string, fallback: string | undefined): string | undefined => {
		const translated = translate(key);
		return translated && translated !== key ? translated : fallback;
	};

	const descriptor = computed(() => buildRibbonGallery(gallery(), host.context.value));
	const disabled = computed(
		() =>
			host.editable?.value === false ||
			descriptor.value.disabled ||
			!galleryHasItems(descriptor.value),
	);
	return {
		descriptor,
		disabled,
		inlineItems: computed(() => inlineGalleryItems(descriptor.value)),
		title: computed(
			() => orFallback(descriptor.value.labelKey, descriptor.value.label) ?? descriptor.value.label,
		),
		itemLabel: (item) => galleryItemLabel(item, translate),
		sectionTitle: (section) =>
			section.titleKey ? orFallback(section.titleKey, section.title) : section.title,
		pick: (itemId, action) => {
			if (disabled.value) {
				return null;
			}
			const result = applyRibbonGalleryItem(gallery(), itemId, host.context.value, action);
			if (!result) {
				return null;
			}
			if (result.kind === 'command') {
				host.endPreview?.();
			} else {
				host.dispatch(result);
			}
			return result;
		},
		preview: (itemId) => {
			host.preview?.(applyRibbonGalleryItem(gallery(), itemId, host.context.value));
		},
		endPreview: () => host.endPreview?.(),
		translate,
	};
}
