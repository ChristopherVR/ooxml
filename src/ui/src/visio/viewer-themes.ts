import { canEditDiagram } from './controller';
import { VISIO_BUILT_IN_THEMES, type VisioPageThemeEdit } from 'ooxml-core/visio';
import type { OfficeUiGallery } from '../ribbon/gallery';
import type { ViewerController, ViewerState } from './controller';
import type { VisioPage } from 'ooxml-core/visio';
import type { VisioRibbonAction } from './ribbon-action';
import {
	NO_THEME_VARIANTS,
	THEMES_HINT,
	VARIANT_EMBELLISHMENTS,
	VARIANT_SCHEME_MENUS,
	schemeItemId,
	VARIANTS_HINT,
	themesGalleryState,
	variantsGalleryState,
} from './ribbon-themes';

type ThemeAction = Extract<VisioRibbonAction, { type: 'page-theme' }>;

/** Design > Themes and Variants: per-page theme edits and their gallery state. */
export class ViewerThemes {
	constructor(
		private readonly root: ShadowRoot,
		private readonly controller: ViewerController,
		private readonly edit: (run: () => Promise<void>, success: string) => void,
	) {}

	#reason(state: ViewerState): string {
		return state.loading
			? 'Opening diagram.'
			: state.edit.busy
				? 'Updating diagram.'
				: !canEditDiagram(state)
					? 'Open a .vsdx file to apply a theme.'
					: !state.document?.pages[state.pageIndex]
						? 'The drawing has no page.'
						: '';
	}

	run(action: ThemeAction): void {
		const state = this.controller.state;
		const page = state.document?.pages[state.pageIndex];
		if (!page || this.#reason(state)) return;
		if (action.embellishment !== undefined) {
			const level = VARIANT_EMBELLISHMENTS.find(([key]) => key === action.embellishment);
			if (!level) return;
			const edit: VisioPageThemeEdit = {
				type: 'set-page-theme',
				pageId: page.id,
				embellishment: level[0],
			};
			this.edit(
				() => this.controller.applyEdits([edit]),
				`Applied ${level[1]} embellishment to ${page.name}.`,
			);
			return;
		}
		if (action.variant !== undefined && !page.theme) return;
		const scheme = VARIANT_SCHEME_MENUS.find(([key]) => action[key] !== undefined)?.[0];
		const edit: VisioPageThemeEdit = {
			type: 'set-page-theme',
			pageId: page.id,
			...(action.theme === undefined ? {} : { theme: action.theme }),
			...(scheme ? { [scheme]: action[scheme] } : {}),
			...(action.variant === undefined ? {} : { variant: action.variant }),
		};
		const name =
			action.theme === 'none'
				? 'No Theme'
				: VISIO_BUILT_IN_THEMES.find(
						(theme) => theme.id === (action.theme ?? (scheme && action[scheme])),
					)?.name;
		this.edit(
			() => this.controller.applyEdits([edit]),
			name
				? `Applied ${name}${scheme ? ` ${scheme}` : ''} to ${page.name}.`
				: `Applied variant ${(action.variant ?? 0) + 1} to ${page.name}.`,
		);
	}

	/** Mark the built-in theme behind each scheme of the page, and gate what needs a theme. */
	#schemes(page: VisioPage | undefined, reason: string): void {
		for (const [scheme, menuId, label] of VARIANT_SCHEME_MENUS) {
			const menu = this.root.querySelector<HTMLElement>(`[data-menu="${menuId}"]`);
			if (!menu) continue;
			// Effect and connector styles are formats of the page's theme: without one they draw nothing.
			const needsTheme = scheme === 'effects' || scheme === 'connectors';
			const why = reason || (needsTheme && !page?.theme ? NO_THEME_VARIANTS : '');
			menu.toggleAttribute('disabled', !!why);
			if (why) menu.title = `${label}: ${why}`;
			else menu.removeAttribute('title');
			const current = page?.themeSchemes?.[scheme];
			const checked = current && schemeItemId(scheme, current);
			for (const item of menu.querySelectorAll<HTMLElement>('office-ui-menu-item'))
				item.setAttribute('checked', String(item.getAttribute('command') === checked));
		}
	}

	render(state: ViewerState): void {
		const page = state.document?.pages[state.pageIndex];
		const reason = this.#reason(state);
		const themes = this.root.querySelector<OfficeUiGallery>(
			'office-ui-gallery[data-menu="themes"]',
		);
		if (themes) {
			themes.state = themesGalleryState(!!reason, page?.theme);
			themes.toggleAttribute('disabled', !!reason);
			themes.title = reason ? `Themes: ${reason}` : THEMES_HINT;
		}
		const options = this.root.querySelector<HTMLElement>('[data-menu="variant-options"]');
		if (options) options.toggleAttribute('disabled', !!reason);
		this.#schemes(page, reason);
		const variants = this.root.querySelector<OfficeUiGallery>(
			'office-ui-gallery[data-menu="variants"]',
		);
		if (variants) {
			const why = reason || (page?.theme ? '' : NO_THEME_VARIANTS);
			variants.state = variantsGalleryState(!!why, page?.theme);
			variants.toggleAttribute('disabled', !!why);
			variants.title = why ? `Variants: ${why}` : VARIANTS_HINT;
		}
	}
}
