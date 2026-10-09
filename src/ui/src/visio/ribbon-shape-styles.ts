import {
	VISIO_QUICK_STYLE_COLORS,
	VISIO_SHADOW_PRESETS,
	visioFallbackQuickStyle,
	type VisioQuickStyleColor,
	type VisioShadowPreset,
} from 'ooxml-core/visio';
import type { OfficeGalleryItem, OfficeGalleryState, OfficeUiGallery } from '../ribbon/gallery';
import { emitRibbonAction } from './ribbon-action';
import type { CommandSpec } from './ribbon-parts';

const COLOR_NAMES: Readonly<Record<VisioQuickStyleColor, string>> = {
	0: 'Dark',
	2: 'Accent 1',
	3: 'Accent 2',
	4: 'Accent 3',
	5: 'Accent 4',
	6: 'Accent 5',
	7: 'Accent 6',
};
const MATRIX_NAMES = ['Subtle', 'Refined', 'Balanced', 'Moderate', 'Focused', 'Intense'];
const SHADOW_NAMES: Readonly<Record<VisioShadowPreset, string>> = {
	none: 'No Shadow',
	'bottom-right': 'Offset: Bottom Right',
	bottom: 'Offset: Bottom',
	'bottom-left': 'Offset: Bottom Left',
	right: 'Offset: Right',
	center: 'Offset: Center',
	left: 'Offset: Left',
	'top-right': 'Offset: Top Right',
	top: 'Offset: Top',
	'top-left': 'Offset: Top Left',
};
export const QUICK_STYLES_HINT =
	"Theme styles from the drawing's theme. A drawing without a theme gets the fixed Office colours shown here.";

/** Stable gallery item id for a colour slot and style matrix. */
export const quickStyleItemId = (color: VisioQuickStyleColor, matrix: number) =>
	`quick-style-${color}-${matrix}`;

function preview(color: VisioQuickStyleColor, matrix: number): string {
	const { fill, line, font } = visioFallbackQuickStyle({ color, matrix });
	return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 40 28"><rect x="2" y="2" width="36" height="24" rx="2" fill="${fill}" stroke="${line}" stroke-width="2"/><text x="20" y="18.5" font-size="10" font-family="Segoe UI, sans-serif" text-anchor="middle" fill="${font}">Abc</text></svg>`;
}

/** The gallery's items: one row per style matrix, one column per theme colour, as in Visio. */
export function quickStyleItems(): OfficeGalleryItem[] {
	return MATRIX_NAMES.flatMap((name, row) =>
		VISIO_QUICK_STYLE_COLORS.map((color) => ({
			id: quickStyleItemId(color, row + 1),
			label: `${COLOR_NAMES[color]}, ${name}`,
			preview: preview(color, row + 1),
		})),
	);
}

export function quickStyleGalleryState(disabled: boolean): OfficeGalleryState {
	return {
		id: 'quick-styles',
		label: 'Quick Styles',
		disabled,
		sections: [
			{
				title: 'Theme Styles',
				columns: 7,
				tileWidth: 40,
				tileHeight: 28,
				items: quickStyleItems(),
			},
		],
	};
}

/** Home > Shape Styles > Quick Styles: the shared Office gallery raising typed format actions. */
export function quickStyleGallery(doc: Document): HTMLElement {
	const gallery = doc.createElement('office-ui-gallery') as OfficeUiGallery;
	gallery.dataset.menu = 'quick-styles';
	gallery.setAttribute('icon', 'quickStyles');
	gallery.setAttribute('label', 'Quick Styles');
	gallery.title = QUICK_STYLES_HINT;
	gallery.state = quickStyleGalleryState(true);
	gallery.toggleAttribute('disabled', true);
	gallery.addEventListener('office-gallery-pick', (event) => {
		event.stopPropagation();
		const match = /^quick-style-(\d+)-(\d+)$/.exec(
			(event as CustomEvent<{ itemId: string }>).detail.itemId,
		);
		if (!match || gallery.hasAttribute('disabled')) return;
		const color = Number(match[1]) as VisioQuickStyleColor;
		emitRibbonAction(gallery, {
			type: 'shape-format',
			patch: { quickStyle: { color, matrix: Number(match[2]) } },
		});
	});
	// KeyTips activate a control by clicking its host; open the dropdown for them.
	gallery.addEventListener('click', (event) => {
		if (event.target === gallery && !gallery.hasAttribute('disabled')) gallery.open = !gallery.open;
	});
	return gallery;
}

const EFFECT_REASONS: readonly [id: string, label: string, reason: string][] = [
	['reflection', 'Reflection', 'Reflection cells are not written or rendered yet.'],
	['glow', 'Glow', 'Glow cells are not written or rendered yet.'],
	['soft-edges', 'Soft Edges', 'Soft edge cells are not written or rendered yet.'],
	['bevel', 'Bevel', 'Bevel needs 3-D lighting the SVG renderer does not draw.'],
	['rotation-3d', '3-D Rotation', '3-D rotation needs a 3-D renderer.'],
];

/** Home > Shape Styles > Effects: outer shadow presets; the other effects stay disabled. */
export function effectsOptions(): CommandSpec[] {
	return [
		{
			id: 'shadow',
			label: 'Shadow',
			items: VISIO_SHADOW_PRESETS.map((preset) => ({
				id: `shadow-${preset}`,
				label: SHADOW_NAMES[preset],
				action: { type: 'shape-format' as const, patch: { shadow: preset } },
				checked: false,
			})),
		},
		...EFFECT_REASONS.map(([id, label, unsupported]) => ({ id, label, unsupported })),
	];
}
