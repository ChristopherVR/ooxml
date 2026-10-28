import { artisticPictureNodes } from './picture-artistic-preview';
import { standardPictureFilter } from './picture-standard-filter';
import { pictureLuminanceTransfer, type DrawingPictureEffect } from 'ooxml-core/drawingml';
const NS = 'http://www.w3.org/2000/svg';
let serial = 0;
/** Ordered supported picture effects. Unknown operations remain in the model and are reported. */
export function pictureEffectFilter(
	doc: Document,
	defs: SVGDefsElement,
	effects: readonly DrawingPictureEffect[],
	scheme: Readonly<Record<string, string>> = {},
): { filter?: string; unapplied: string[]; approximated?: string[] } {
	const filter = doc.createElementNS(NS, 'filter');
	filter.id = `office-picture-effects-${++serial}`;
	filter.setAttribute('color-interpolation-filters', 'sRGB');
	const unapplied: string[] = [];
	const approximated: string[] = [];
	let previous = 'SourceGraphic',
		step = 0;
	for (const effect of effects) {
		if (effect.kind === 'unsupported') {
			unapplied.push(effect.name);
			continue;
		}
		if (effect.kind === 'artistic') {
			if (effect.baked) continue;
			const nodes = artisticPictureNodes(doc, effect.name, effect.radius);
			if (!nodes) {
				unapplied.push(effect.name);
				continue;
			}
			approximated.push(effect.name);
			const input = previous;
			const names = new Map(
				nodes.flatMap((node) =>
					node.hasAttribute('result')
						? [[node.getAttribute('result')!, `${filter.id}-art${step++}`] as const]
						: [],
				),
			);
			for (const node of nodes) {
				for (const attribute of ['in', 'in2']) {
					const ref = node.getAttribute(attribute);
					if (ref === 'SourceGraphic') node.setAttribute(attribute, input);
					else if (ref && names.has(ref)) node.setAttribute(attribute, names.get(ref)!);
				}
				const result = node.getAttribute('result');
				if (result && names.has(result)) node.setAttribute('result', names.get(result)!);
				else node.setAttribute('result', `${filter.id}-r${step++}`);
				if (!node.hasAttribute('in') && !['feTurbulence', 'feFlood'].includes(node.localName))
					node.setAttribute('in', previous);
				previous = node.getAttribute('result')!;
				filter.append(node);
			}
			continue;
		}
		if (effect.kind === 'standard') {
			const result = standardPictureFilter(doc, effect, scheme, `${filter.id}-${step++}`);
			if (result.unapplied) unapplied.push(result.unapplied);
			if (['hsl', 'tint', 'clrChange'].includes(effect.name)) approximated.push(effect.name);
			const input = previous;
			for (const item of result.nodes) {
				if (item.getAttribute('in') === 'PICTURE_INPUT') item.setAttribute('in', input);
				else if (!item.hasAttribute('in') && item.localName !== 'feFlood')
					item.setAttribute('in', previous);
				if (!item.hasAttribute('result')) item.setAttribute('result', `${filter.id}-r${step++}`);
				previous = item.getAttribute('result')!;
				filter.append(item);
			}
			continue;
		}
		if (effect.kind === 'grayscale') {
			const node = doc.createElementNS(NS, 'feColorMatrix');
			node.setAttribute('type', 'saturate');
			node.setAttribute('values', '0');
			node.setAttribute('in', previous);
			node.setAttribute('result', `${filter.id}-r${step++}`);
			previous = node.getAttribute('result')!;
			filter.append(node);
		} else {
			const transfer = doc.createElementNS(NS, 'feComponentTransfer');
			if (effect.kind === 'alphaModFix') {
				const channel = doc.createElementNS(NS, 'feFuncA');
				channel.setAttribute('type', 'linear');
				channel.setAttribute('slope', String(effect.amount));
				transfer.append(channel);
			} else {
				const { slope, intercept } = pictureLuminanceTransfer(effect.brightness, effect.contrast);
				for (const name of ['R', 'G', 'B']) {
					const channel = doc.createElementNS(NS, `feFunc${name}`);
					channel.setAttribute('type', 'linear');
					channel.setAttribute('slope', String(slope));
					channel.setAttribute('intercept', String(intercept));
					transfer.append(channel);
				}
			}
			transfer.setAttribute('in', previous);
			transfer.setAttribute('result', `${filter.id}-r${step++}`);
			previous = transfer.getAttribute('result')!;
			filter.append(transfer);
		}
	}
	if (!filter.childNodes.length) return { unapplied };
	defs.append(filter);
	return {
		filter: `url(#${filter.id})`,
		unapplied,
		...(approximated.length ? { approximated } : {}),
	};
}
