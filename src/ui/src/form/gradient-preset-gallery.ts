import {
	OFFICE_GRADIENT_PRESETS,
	officeGradientPresetFill,
	officeGradientPresetId,
	resolveDrawingColor,
	type DiagramFill,
} from 'ooxml-core/diagram';
import { resolveChartGradient } from 'ooxml-core/chart';
import {
	defineGallery,
	type OfficeUiGallery,
	type OfficeGalleryPickEvent,
} from '../ribbon/gallery';
import { gradientGalleryPreview } from './gradient-gallery-preview';

interface PresetOptions {
	fill: DiagramFill;
	disabled: boolean;
	label: string;
	translate(label: string): string;
	onPick(id: number): void;
}
let sequence = 0;

/** Native preset paint with the shared Office popup and keyboard navigation. */
export function createGradientPresetGallery(doc: Document) {
	const prefix = `office-gradient-preset-${++sequence}`;
	defineGallery(doc.defaultView?.customElements);
	const element = doc.createElement('office-ui-gallery') as OfficeUiGallery;
	let current: PresetOptions | undefined;
	element.addEventListener('office-gallery-pick', (event) => {
		const id = (event as OfficeGalleryPickEvent).detail.itemId;
		const preset = OFFICE_GRADIENT_PRESETS.find((preset) => String(preset.id) === id);
		if (preset && current && !current.disabled) current.onPick(preset.id);
	});
	return {
		element,
		close: () => {
			element.open = false;
		},
		update(options: PresetOptions) {
			current = options;
			if (options.disabled) element.open = false;
			const selected = officeGradientPresetId(options.fill);
			element.state = {
				id: 'gradient-presets',
				label: options.label,
				moreLabel: options.label,
				disabled: options.disabled,
				sections: [
					{
						columns: 6,
						tileWidth: 40,
						tileHeight: 40,
						items: OFFICE_GRADIENT_PRESETS.map((preset) => ({
							id: String(preset.id),
							label: options.translate(preset.label),
							applied: preset.id === selected,
							preview: gradientGalleryPreview(
								`${prefix}-${preset.id}`,
								resolveChartGradient(officeGradientPresetFill(preset.id), (color) =>
									resolveDrawingColor(color),
								),
							),
						})),
					},
				],
			};
		},
	};
}
