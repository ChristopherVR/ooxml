import { OFFICE_GRADIENT_PRESET_DATA } from './gradient-preset-data';
import { NS } from '../xml';
import type { DiagramFill } from './types';

type Gradient = Extract<DiagramFill, { kind: 'gradient' }>;

/** The 24 standard Office named presets, with native horizontal style/variant 1. */
export const OFFICE_GRADIENT_PRESETS: ReadonlyArray<{
	readonly id: number;
	readonly label: string;
}> = OFFICE_GRADIENT_PRESET_DATA.map(([label], index) => ({
	id: index + 1,
	label,
}));

/** Fresh native DrawingML paint; applying a preset replaces prior stops and geometry. */
export function officeGradientPresetFill(id: number): Gradient {
	const data = Number.isInteger(id) && OFFICE_GRADIENT_PRESET_DATA[id - 1];
	if (!data) throw new RangeError(`Unknown Office gradient preset ${id}`);
	return {
		kind: 'gradient',
		angle: 90,
		scaled: true,
		sourceXml: `<a:gradFill xmlns:a="${NS.a}" flip="none" rotWithShape="1"><a:tileRect/></a:gradFill>`,
		stops: data[1].map(([position, color]) => ({
			position: position / 1000,
			color: { kind: 'srgb', value: color, transforms: [] },
		})),
	};
}

/** Recognize exact preset values; customized or unsupported paint has no preset identity. */
export function officeGradientPresetId(fill: DiagramFill): number | undefined {
	if (fill.kind !== 'gradient' || fill.path || fill.angle !== 90 || fill.scaled !== true)
		return undefined;
	const index = OFFICE_GRADIENT_PRESET_DATA.findIndex(
		([, stops]) =>
			stops.length === fill.stops.length &&
			stops.every(([position, color], index) => {
				const stop = fill.stops[index];
				return (
					stop?.position === position / 1000 &&
					stop.color.kind === 'srgb' &&
					stop.color.value.toUpperCase() === color &&
					stop.color.transforms.length === 0
				);
			}),
	);
	return index < 0 ? undefined : index + 1;
}
