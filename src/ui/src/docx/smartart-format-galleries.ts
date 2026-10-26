import {
	DIAGRAM_COLOR_ENTRIES,
	DIAGRAM_STYLE_ENTRIES,
	type DiagramColorScheme,
	type DiagramStyleIntensity,
} from 'ooxml-core/diagram';
import { smartArtColorsGalleryState, smartArtStylesGalleryState } from '../smartart';
import {
	defineGallery,
	type OfficeUiGallery,
	type OfficeGalleryPickEvent,
} from '../ribbon/gallery';
import { emit } from './events';
import { normalizeEditorLocale, translate, type LocalizationKey } from './localization';

export function smartArtFormatGallery(kind: 'colors' | 'style'): HTMLElement {
	defineGallery();
	const gallery = document.createElement('office-ui-gallery') as OfficeUiGallery;
	gallery.dataset.smartartFormat = kind;
	gallery.addEventListener('office-gallery-pick', (event) => {
		const value = (event as OfficeGalleryPickEvent).detail.itemId;
		if (kind === 'colors' && DIAGRAM_COLOR_ENTRIES.some((entry) => entry.id === value))
			emit(gallery, 'ribbon-action', {
				type: 'smartArtFormat',
				change: { kind, value: value as DiagramColorScheme },
			});
		if (kind === 'style' && DIAGRAM_STYLE_ENTRIES.some((entry) => entry.id === value))
			emit(gallery, 'ribbon-action', {
				type: 'smartArtFormat',
				change: { kind, value: value as DiagramStyleIntensity },
			});
	});
	return gallery;
}

export function syncSmartArtFormatGalleries(
	root: HTMLElement,
	disabled?: boolean,
	colorsId?: string,
	styleId?: string,
): void {
	const locale = normalizeEditorLocale(root.dataset.editorLocale);
	const t = (key: string) => translate(locale, key as LocalizationKey);
	for (const gallery of root.querySelectorAll<OfficeUiGallery>('[data-smartart-format]')) {
		const state =
			gallery.dataset.smartartFormat === 'colors'
				? smartArtColorsGalleryState(
						{},
						colorsId?.split('/').at(-1) as DiagramColorScheme | undefined,
						t,
					)
				: smartArtStylesGalleryState(
						undefined,
						styleId?.split('/').at(-1) as DiagramStyleIntensity | undefined,
						t,
					);
		gallery.state = {
			...state,
			moreLabel: state.label,
			disabled: disabled ?? gallery.state?.disabled ?? false,
		};
	}
}
