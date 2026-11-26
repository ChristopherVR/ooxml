/** Contextual groups and the host customization IDs for their controls. */
export const RIBBON_CATALOG_MEDIA_TABS = {
	videoFormat: {
		poster: {
			label: 'Poster Frame',
			controls: { current: 'Current Frame', file: 'Image from File', reset: 'Reset Poster Frame' },
		},
		adjust: { label: 'Adjust', controls: { corrections: 'Corrections', color: 'Color' } },
		pictureStyles: {
			label: 'Video Styles',
			controls: {
				gallery: 'Video Styles',
				shape: 'Video Shape',
				border: 'Video Border',
				effects: 'Video Effects',
			},
		},
		arrange: { label: 'Arrange', controls: { rotateRight: 'Rotate Right 90°' } },
		size: { label: 'Size', controls: { width: 'Width', height: 'Height', crop: 'Crop' } },
	},
	audioFormat: {
		adjust: { label: 'Adjust', controls: { corrections: 'Corrections', color: 'Color' } },
		pictureStyles: {
			label: 'Picture Styles',
			controls: {
				gallery: 'Picture Styles',
				shape: 'Shape',
				border: 'Border',
				effects: 'Picture Effects',
			},
		},
		arrange: { label: 'Arrange', controls: { rotateRight: 'Rotate Right 90°' } },
		size: { label: 'Size', controls: { width: 'Width', height: 'Height' } },
	},
	playback: {
		bookmarks: { label: 'Bookmarks', controls: { add: 'Add Bookmark', remove: 'Remove Bookmark' } },
		editing: {
			label: 'Editing',
			controls: { trim: 'Trim', fadeIn: 'Fade In', fadeOut: 'Fade Out' },
		},
		options: {
			label: 'Playback',
			controls: {
				volume: 'Volume',
				start: 'Start',
				fullScreen: 'Play Full Screen',
				hide: 'Hide While Not Playing',
				loop: 'Loop until Stopped',
				rewind: 'Rewind after Playing',
				across: 'Play Across Slides',
			},
		},
	},
	smartArtFormat: {
		shapes: {
			label: 'Shapes',
			controls: {
				shape: 'Change Shape',
				larger: 'Larger',
				smaller: 'Smaller',
			},
		},
		shapeStyles: {
			label: 'Shape Styles',
			controls: { gallery: 'Shape Styles', fill: 'Shape Fill', border: 'Shape Outline' },
		},
		wordArtStyles: {
			label: 'WordArt Styles',
			controls: { gallery: 'WordArt Styles', color: 'Text Fill', bold: 'Bold', italic: 'Italic' },
		},
		arrange: { label: 'Arrange', controls: {} },
		size: { label: 'Size', controls: { width: 'Width', height: 'Height' } },
	},
} as const;
