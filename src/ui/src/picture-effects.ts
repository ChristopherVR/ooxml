import { pictureLuminanceTransfer, type DrawingPictureEffect } from 'ooxml-core/drawingml';
const NS = 'http://www.w3.org/2000/svg';
let serial = 0;
/** Ordered supported picture effects. Unknown operations remain in the model and are reported. */
export function pictureEffectFilter(
	doc: Document,
	defs: SVGDefsElement,
	effects: readonly DrawingPictureEffect[],
): { filter?: string; unapplied: string[] } {
	const filter = doc.createElementNS(NS, 'filter');
	filter.id = `office-picture-effects-${++serial}`;
	filter.setAttribute('color-interpolation-filters', 'sRGB');
	const unapplied: string[] = [];
	for (const effect of effects) {
		if (effect.kind === 'unsupported') {
			if (effect.name !== 'extLst') unapplied.push(effect.name);
			continue;
		}
		if (effect.kind === 'grayscale') {
			const node = doc.createElementNS(NS, 'feColorMatrix');
			node.setAttribute('type', 'saturate');
			node.setAttribute('values', '0');
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
			filter.append(transfer);
		}
	}
	if (!filter.childNodes.length) return { unapplied };
	defs.append(filter);
	return { filter: `url(#${filter.id})`, unapplied };
}
