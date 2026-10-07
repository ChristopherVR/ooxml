import { withChartGradientDirection, type ChartGradientFill } from 'ooxml-core/chart';
import {
	RECT_GRADIENT_DIRECTIONS,
	rectGradientDirection,
	type RectGradientDirection,
} from 'ooxml-core/diagram';
import { gradientGalleryPreview } from './gradient-gallery-preview';
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
	onRectPick?(direction: RectGradientDirection): void;
	onPathPick?(direction: RectGradientDirection): void;
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
		if (['rect', 'circle', 'shape'].includes(current?.gradient.path ?? '')) {
			const direction = RECT_GRADIENT_DIRECTIONS.find((direction) => direction.id === id);
			if (direction && current && !current.disabled) {
				if (current.onPathPick) current.onPathPick(direction.id);
				else if (current.gradient.path === 'rect') current.onRectPick?.(direction.id);
			}
			return;
		}
		const direction = DIRECTIONS.find(([angle]) => String(angle) === id);
		if (direction && current && !current.disabled) current.onPick(direction[0]);
	});
	const update = (options: DirectionOptions) => {
		current = options;
		if (options.disabled) element.open = false;
		const rectangular = ['rect', 'circle', 'shape'].includes(options.gradient.path ?? '');
		const items = rectangular
			? RECT_GRADIENT_DIRECTIONS.map(({ id, label }) => ({
					id,
					label: options.translate(label),
					applied: rectGradientDirection(options.gradient.fillToRect) === id,
					preview: gradientGalleryPreview(
						`${prefix}-${id}`,
						withChartGradientDirection(options.gradient, id),
					),
				}))
			: DIRECTIONS.map(([angle, label]) => ({
					id: String(angle),
					label: options.translate(label),
					applied: options.gradient.angle === angle,
					preview: gradientGalleryPreview(`${prefix}-${angle}`, {
						...options.gradient,
						type: 'linear',
						angle,
					}),
				}));
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
					items,
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
