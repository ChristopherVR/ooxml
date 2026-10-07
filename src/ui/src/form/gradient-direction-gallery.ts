import { buildChartGradientDef, type ChartGradientFill } from 'ooxml-core/chart';
import {
	defineGallery,
	type OfficeUiGallery,
	type OfficeGalleryPickEvent,
} from '../ribbon/gallery';

const DIRECTIONS = [
	[0, 'Linear Right'],
	[45, 'Linear Diagonal - Bottom Right'],
	[90, 'Linear Down'],
	[135, 'Linear Diagonal - Bottom Left'],
	[180, 'Linear Left'],
	[225, 'Linear Diagonal - Top Left'],
	[270, 'Linear Up'],
	[315, 'Linear Diagonal - Top Right'],
] as const;

interface DirectionOptions {
	gradient: ChartGradientFill;
	disabled: boolean;
	label: string;
	translate(label: string): string;
	onPick(angle: number): void;
}
let sequence = 0;

/** Shared Office gallery and chart painter provide previews, focus and popup behavior. */
export function createGradientDirectionGallery(doc: Document) {
	const prefix = `office-gradient-direction-${++sequence}`;
	defineGallery(doc.defaultView?.customElements);
	const element = doc.createElement('office-ui-gallery') as OfficeUiGallery;
	let current: DirectionOptions | undefined;
	element.addEventListener('office-gallery-pick', (event) => {
		const id = (event as OfficeGalleryPickEvent).detail.itemId;
		const direction = DIRECTIONS.find(([angle]) => String(angle) === id);
		if (direction && current && !current.disabled) current.onPick(direction[0]);
	});
	const update = (options: DirectionOptions) => {
		current = options;
		if (options.disabled) element.open = false;
		element.state = {
			id: 'gradient-direction',
			label: options.label,
			moreLabel: options.label,
			disabled: options.disabled,
			sections: [
				{
					columns: 4,
					tileWidth: 44,
					tileHeight: 44,
					items: DIRECTIONS.map(([angle, label]) => {
						const def = buildChartGradientDef(`${prefix}-${angle}`, {
							...options.gradient,
							type: 'linear',
							angle,
						});
						if (def.kind !== 'linearGradient') throw new Error('Expected linear gradient');
						const stops = def.stops
							.map(
								(stop) =>
									`<stop offset="${stop.offset}" stop-color="${stop.color}" stop-opacity="${stop.opacity ?? 1}"/>`,
							)
							.join('');
						return {
							id: String(angle),
							label: options.translate(label),
							applied: options.gradient.angle === angle,
							preview: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 40 40"><defs><linearGradient id="${def.id}" x1="${def.x1}" y1="${def.y1}" x2="${def.x2}" y2="${def.y2}">${stops}</linearGradient></defs><rect x="1" y="1" width="38" height="38" fill="url(#${def.id})" stroke="#999"/></svg>`,
						};
					}),
				},
			],
		};
	};
	return {
		element,
		update,
		close: () => {
			element.open = false;
		},
	};
}
