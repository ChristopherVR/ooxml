import type { PptxElement } from 'ooxml-core/pptx';

import type { RibbonGalleryModule } from './gallery-module';
import { adjustGalleryModule, sameHex } from './picture-adjust-gallery';
import type { AdjustPreset } from './picture-adjust-gallery';

const STANDARD = [
	['black', '#000000'],
	['white', '#ffffff'],
	['red', '#ff0000'],
	['green', '#00b050'],
	['blue', '#0070c0'],
] as const;

function eligible(element: PptxElement | null): boolean {
	return (
		(element?.type === 'image' || element?.type === 'picture') &&
		Boolean(element.svgData || element.svgPath)
	);
}

const gallery = adjustGalleryModule(
	'graphicsFill',
	'pptx.gallery.graphicsFill',
	'Graphics Fill',
	(ctx) => {
		const swatch = (id: string, labelKey: string, label: string, color: string): AdjustPreset => ({
			id,
			labelKey,
			label,
			changes: () => ({ clrRepl: { color } }),
			applied: (effects) => sameHex(effects.clrRepl?.color, color),
			preview: { duotone: [color, color] },
		});
		const presets = STANDARD.map(([name, color]) =>
			swatch(name, `pptx.colorPicker.standard.${name}`, name, color),
		);
		const theme = Array.from({ length: 6 }, (_, i) => {
			const id = `accent${i + 1}`;
			const color = ctx.themeColorMap?.[id];
			return color && /^#[a-f\d]{6}$/i.test(color)
				? [swatch(id, `pptx.themeColor.${id}`, `Accent ${i + 1}`, color)]
				: [];
		}).flat();
		return [
			{
				id: 'fill',
				columns: 6,
				presets: [
					{
						id: 'original',
						labelKey: 'pptx.gallery.pictureColor.recolorNone',
						label: 'No Recolor',
						changes: () => ({ clrRepl: undefined }),
						applied: (effects) => !effects.clrRepl,
						preview: {},
					},
					...presets,
					...theme,
				],
			},
		];
	},
);

export const GRAPHICS_FILL_GALLERY: RibbonGalleryModule = {
	build(ctx) {
		const view = gallery.build(ctx);
		return eligible(ctx.element) ? view : { ...view, disabled: true, sections: [] };
	},
	apply(id, ctx) {
		return eligible(ctx.element) ? gallery.apply(id, ctx) : null;
	},
};
