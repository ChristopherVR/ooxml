import type { RibbonControlId, RibbonGroupId } from './customization';

export const DESIGN_RIBBON_GROUPS: readonly {
	id: RibbonGroupId;
	labelKey: string;
}[] = [
	{ id: 'design.themes', labelKey: 'pptx.ribbon.groupThemes' },
	{ id: 'design.variants', labelKey: 'pptx.ribbon.groupVariants' },
	{ id: 'design.customize', labelKey: 'pptx.ribbon.groupCustomize' },
	{ id: 'design.designer', labelKey: 'pptx.ribbon.groupDesigner' },
];

export interface DesignRibbonState {
	editable: boolean;
	galleryOpen?: boolean;
	editorOpen?: boolean;
	backgroundOpen?: boolean;
	hasSlideSize?: boolean;
	hasBackground?: boolean;
}

export function designCommandState(id: RibbonControlId, state: DesignRibbonState) {
	const slideSize = id === 'design.customize.slideSize';
	const background = id === 'design.customize.formatBackground';
	// Designer needs an online service the viewer does not have: always listed, never enabled.
	const designer = id === 'design.designer.designIdeas';
	const active =
		id === 'design.themes.browseThemes'
			? state.galleryOpen
			: id === 'design.themes.editTheme'
				? state.editorOpen
				: background
					? state.backgroundOpen
					: false;
	return {
		disabled: designer || (!slideSize && !state.editable),
		active: Boolean(active),
		expanded: slideSize ? undefined : Boolean(active),
		hidden:
			(slideSize && state.hasSlideSize === false) || (background && state.hasBackground === false),
	};
}

/** Presentation only: adapters retain their gallery, editor and inspector actions. */
export const DESIGN_RIBBON_COMMANDS: readonly {
	id: RibbonControlId;
	labelKey: string;
	titleKey: string;
	icon: string;
}[] = [
	{
		id: 'design.themes.browseThemes',
		labelKey: 'pptx.ribbon.browseThemes',
		titleKey: 'pptx.ribbon.browseThemesTitle',
		icon: 'palette',
	},
	{
		id: 'design.themes.editTheme',
		labelKey: 'pptx.ribbon.editTheme',
		titleKey: 'pptx.design.editThemeTooltip',
		icon: 'pencil',
	},
	{
		id: 'design.customize.slideSize',
		labelKey: 'pptx.ribbon.slideSize',
		titleKey: 'pptx.ribbon.slideSizeTitle',
		icon: 'monitor',
	},
	{
		id: 'design.customize.formatBackground',
		labelKey: 'pptx.ribbon.formatBackground',
		titleKey: 'pptx.ribbon.formatBackgroundTitle',
		icon: 'paint',
	},
	{
		id: 'design.designer.designIdeas',
		labelKey: 'pptx.ribbon.designIdeas',
		titleKey: 'pptx.design.designIdeasTooltip',
		icon: 'lightbulb',
	},
];
