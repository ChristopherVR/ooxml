import { findDiagramNativeColor, findDiagramNativeStyle } from 'ooxml-core/diagram';
import { smartArtColorsGalleryState, smartArtStylesGalleryState } from '../smartart';
import {
	defineGallery,
	type OfficeUiGallery,
	type OfficeGalleryPickEvent,
} from '../ribbon/gallery';
import { emit } from './events';
import { normalizeEditorLocale, translate, type LocalizationKey } from './localization';

// The built-in definitions are in the separate `diagram/builtin` entry, loaded once the galleries
// are enabled (a SmartArt is selected), so a pick normally applies at once (a pick during the
// load waits for it). Building the ribbon never loads them.
type Builtin = typeof import('ooxml-core/diagram/builtin');
let loaded: Builtin | undefined;
let loading: Promise<Builtin> | undefined;
const builtin = (): Promise<Builtin> =>
	(loading ??= import('ooxml-core/diagram/builtin').then((module) => (loaded = module)));

export function smartArtFormatGallery(kind: 'colors' | 'style'): HTMLElement {
	defineGallery();
	const gallery = document.createElement('office-ui-gallery') as OfficeUiGallery;
	gallery.dataset.smartartFormat = kind;
	gallery.addEventListener('office-gallery-pick', (event) => {
		const value = (event as OfficeGalleryPickEvent).detail.itemId;
		const entry = kind === 'colors' ? findDiagramNativeColor(value) : findDiagramNativeStyle(value);
		if (!entry) return;
		const apply = ({ builtinSmartArtFormatChange }: Builtin) =>
			emit(gallery, 'ribbon-action', {
				type: 'smartArtFormat',
				change: builtinSmartArtFormatChange(kind, entry.uniqueId),
			});
		if (loaded) apply(loaded);
		else void builtin().then(apply);
	});
	return gallery;
}

export function syncSmartArtFormatGalleries(
	root: HTMLElement,
	disabled?: boolean,
	colorsId?: string,
	styleId?: string,
): void {
	if (disabled === false && !loaded) void builtin().catch(() => undefined);
	const locale = normalizeEditorLocale(root.dataset.editorLocale);
	const t = (key: string) => translate(locale, key as LocalizationKey);
	for (const gallery of root.querySelectorAll<OfficeUiGallery>('[data-smartart-format]')) {
		const state =
			gallery.dataset.smartartFormat === 'colors'
				? smartArtColorsGalleryState({}, colorsId, t)
				: smartArtStylesGalleryState(undefined, styleId, t);
		gallery.state = {
			...state,
			moreLabel: state.label,
			disabled: disabled ?? gallery.state?.disabled ?? false,
		};
	}
}
