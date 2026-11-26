import type { RibbonGalleryGroupPlacement } from './gallery-placements';

function mediaFormat(tab: 'videoFormat' | 'audioFormat'): RibbonGalleryGroupPlacement[] {
	return [
		{
			group: `${tab}.adjust`,
			labelKey: 'pptx.ribbon.groupAdjust',
			label: 'Adjust',
			galleries: [
				{ gallery: 'pictureCorrections', control: `${tab}.adjust.corrections`, mode: 'dropdown' },
				{ gallery: 'pictureColor', control: `${tab}.adjust.color`, mode: 'dropdown' },
			],
		},
		{
			group: `${tab}.pictureStyles`,
			labelKey: 'pptx.ribbon.groupPictureStyles',
			label: 'Picture Styles',
			contextualStrip: true,
			galleries: [
				{ gallery: 'pictureStyles', control: `${tab}.pictureStyles.gallery`, mode: 'inline' },
				{ gallery: 'cropToShape', control: `${tab}.pictureStyles.shape`, mode: 'dropdown' },
				{ gallery: 'shapeEffects', control: `${tab}.pictureStyles.effects`, mode: 'dropdown' },
			],
		},
		{
			group: `${tab}.arrange`,
			labelKey: 'pptx.ribbon.arrange',
			label: 'Arrange',
			arrange: true,
			contextualStrip: true,
			galleries: [],
		},
		{
			group: `${tab}.size`,
			labelKey: 'pptx.ribbon.groupSize',
			label: 'Size',
			contextualStrip: true,
			galleries: [],
		},
	];
}

export const MEDIA_FORMAT_GROUPS = {
	videoFormat: [
		{
			group: 'videoFormat.poster',
			labelKey: 'pptx.media.posterFrame',
			label: 'Poster Frame',
			contextualStrip: true,
			galleries: [],
		},
		...mediaFormat('videoFormat'),
	],
	audioFormat: mediaFormat('audioFormat'),
	playback: [
		{
			group: 'playback.bookmarks',
			labelKey: 'pptx.media.bookmarks',
			label: 'Bookmarks',
			contextualStrip: true,
			galleries: [],
		},
		{
			group: 'playback.editing',
			labelKey: 'pptx.ribbon.editing',
			label: 'Editing',
			contextualStrip: true,
			galleries: [],
		},
		{
			group: 'playback.options',
			labelKey: 'pptx.media.playback',
			label: 'Playback',
			contextualStrip: true,
			galleries: [],
		},
	],
	smartArtFormat: [
		{
			group: 'smartArtFormat.shapes',
			labelKey: 'pptx.drawing.shapes',
			label: 'Shapes',
			contextualStrip: true,
			galleries: [],
		},
		{
			group: 'smartArtFormat.shapeStyles',
			labelKey: 'pptx.ribbon.groupShapeStyles',
			label: 'Shape Styles',
			contextualStrip: true,
			galleries: [],
		},
		{
			group: 'smartArtFormat.wordArtStyles',
			labelKey: 'pptx.ribbon.groupWordArtStyles',
			label: 'WordArt Styles',
			contextualStrip: true,
			galleries: [],
		},
		{
			group: 'smartArtFormat.arrange',
			labelKey: 'pptx.ribbon.arrange',
			label: 'Arrange',
			arrange: true,
			galleries: [],
		},
		{
			group: 'smartArtFormat.size',
			labelKey: 'pptx.ribbon.groupSize',
			label: 'Size',
			contextualStrip: true,
			galleries: [],
		},
	],
} satisfies Record<string, readonly RibbonGalleryGroupPlacement[]>;
