import {
	buildImageLuminanceTransfer,
	resolveDrawingColor,
	type DrawingStandardPictureEffect,
} from 'ooxml-core/drawingml';
import { hexToRgbChannels } from 'ooxml-core/color';
const NS = 'http://www.w3.org/2000/svg';
const thresholdTable = (n: number) =>
	Array.from({ length: 256 }, (_, i) => (i / 255 >= n ? '1' : '0')).join(' ');
export function standardPictureFilter(
	doc: Document,
	effect: DrawingStandardPictureEffect,
	scheme: Readonly<Record<string, string>> = {},
	prefix = 'picture',
): { nodes: SVGElement[]; unapplied?: string } {
	const nodes: SVGElement[] = [];
	const node = (name: string, attributes: Record<string, string | number> = {}) => {
		const item = doc.createElementNS(NS, name);
		for (const [key, value] of Object.entries(attributes)) item.setAttribute(key, String(value));
		return item;
	};
	const transfer = (channels: readonly string[], attributes: Record<string, string | number>) => {
		const item = node('feComponentTransfer');
		for (const channel of channels) item.append(node(`feFunc${channel}`, attributes));
		nodes.push(item);
	};
	const value = effect.values ?? {};
	const colors = effect.colors?.map((color) => {
		const resolved = resolveDrawingColor(color, { scheme: (name) => scheme[name] });
		const rgb = resolved ? hexToRgbChannels(resolved.hex) : undefined;
		return rgb && resolved
			? { r: rgb.r / 255, g: rgb.g / 255, b: rgb.b / 255, a: resolved.alpha }
			: undefined;
	});
	switch (effect.name) {
		case 'alphaInv':
			transfer(['A'], { type: 'linear', slope: -1, intercept: 1 });
			{
				const c = colors?.[0] ?? { r: 0, g: 0, b: 0, a: 1 };
				nodes.push(
					node('feColorMatrix', {
						type: 'matrix',
						values: `0 0 0 0 ${c.r} 0 0 0 0 ${c.g} 0 0 0 0 ${c.b} 0 0 0 ${c.a} 0`,
					}),
				);
			}
			break;
		case 'alphaCeiling':
			transfer(['A'], { type: 'discrete', tableValues: ['0', ...Array(255).fill('1')].join(' ') });
			break;
		case 'alphaFloor':
			transfer(['A'], { type: 'discrete', tableValues: [...Array(255).fill('0'), '1'].join(' ') });
			break;
		case 'alphaRepl':
			transfer(['A'], { type: 'linear', slope: 0, intercept: value.amount ?? 1 });
			break;
		case 'alphaMod':
			transfer(['A'], { type: 'gamma', amplitude: value.amount ?? 1, exponent: 2, offset: 0 });
			break;
		case 'alphaBiLevel':
			transfer(['A'], {
				type: 'discrete',
				tableValues: thresholdTable(value.threshold ?? 0),
			});
			break;
		case 'biLevel':
			nodes.push(node('feColorMatrix', { type: 'saturate', values: 0 }));
			transfer(['R', 'G', 'B'], {
				type: 'discrete',
				tableValues: thresholdTable(value.threshold ?? 0),
			});
			break;
		case 'blur':
			nodes.push(node('feGaussianBlur', { stdDeviation: (value.radiusEmu ?? 0) / 3 }));
			break;
		case 'hsl':
		case 'tint':
			if (value.hue !== undefined)
				nodes.push(node('feColorMatrix', { type: 'hueRotate', values: value.hue }));
			if (value.saturation !== undefined)
				nodes.push(
					node('feColorMatrix', { type: 'saturate', values: Math.max(0, 1 + value.saturation) }),
				);
			if (value.luminance !== undefined || value.amount !== undefined) {
				const x = buildImageLuminanceTransfer((value.luminance ?? value.amount ?? 0) * 100);
				transfer(['R', 'G', 'B'], { type: 'linear', slope: x.slope, intercept: x.intercept });
			}
			break;
		case 'duotone': {
			const a = colors?.[0],
				b = colors?.[1];
			if (!a || !b) return { nodes: [], unapplied: 'duotone:colour' };
			nodes.push(node('feColorMatrix', { type: 'saturate', values: 0 }));
			const map = node('feComponentTransfer');
			for (const channel of ['R', 'G', 'B'] as const) {
				const key = channel.toLowerCase() as 'r' | 'g' | 'b';
				map.append(
					node(`feFunc${channel}`, { type: 'linear', slope: b[key] - a[key], intercept: a[key] }),
				);
			}
			nodes.push(map);
			if (a.a !== 1 || b.a !== 1) return { nodes, unapplied: 'duotone:colour-alpha' };
			break;
		}
		case 'clrRepl': {
			const c = colors?.[0];
			if (!c) return { nodes: [], unapplied: 'clrRepl:colour' };
			nodes.push(
				node('feColorMatrix', {
					type: 'matrix',
					values: `0 0 0 0 ${c.r} 0 0 0 0 ${c.g} 0 0 0 0 ${c.b} 0 0 0 ${c.a} 0`,
				}),
			);
			break;
		}
		case 'clrChange': {
			const a = colors?.[0],
				b = colors?.[1];
			if (!a || !b) return { nodes: [], unapplied: 'clrChange:colour' };
			if (effect.useAlpha !== false && a.a !== 1)
				return { nodes: [], unapplied: 'clrChange:source-alpha' };
			const opaque = node('feComponentTransfer', {
				in: 'PICTURE_INPUT',
				result: `${prefix}-opaque`,
			});
			opaque.append(node('feFuncA', { type: 'linear', slope: 0, intercept: 1 }));
			nodes.push(opaque);
			const from = node('feFlood', {
				'flood-color': `rgb(${a.r * 255},${a.g * 255},${a.b * 255})`,
				result: `${prefix}-key`,
			});
			nodes.push(from);
			nodes.push(
				node('feBlend', {
					in: `${prefix}-opaque`,
					in2: `${prefix}-key`,
					mode: 'difference',
					result: `${prefix}-distance`,
				}),
			);
			nodes.push(
				node('feColorMatrix', {
					in: `${prefix}-distance`,
					type: 'matrix',
					values: '0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 1 1 1 0 0',
				}),
			);
			transfer(['A'], { type: 'discrete', tableValues: ['1', ...Array(255).fill('0')].join(' ') });
			nodes[nodes.length - 1]!.setAttribute('result', `${prefix}-mask`);
			nodes.push(
				node('feColorMatrix', {
					in: 'PICTURE_INPUT',
					type: 'matrix',
					values: `0 0 0 0 ${b.r} 0 0 0 0 ${b.g} 0 0 0 0 ${b.b} 0 0 0 ${b.a} 0`,
					result: `${prefix}-replacement`,
				}),
			);
			nodes.push(
				node('feComposite', {
					in: `${prefix}-replacement`,
					in2: `${prefix}-mask`,
					operator: 'in',
					result: `${prefix}-matched`,
				}),
			);
			nodes.push(
				node('feComposite', {
					in: 'PICTURE_INPUT',
					in2: `${prefix}-mask`,
					operator: 'out',
					result: `${prefix}-rest`,
				}),
			);
			nodes.push(
				node('feComposite', { in: `${prefix}-matched`, in2: `${prefix}-rest`, operator: 'over' }),
			);
			break;
		}
	}
	return { nodes };
}
