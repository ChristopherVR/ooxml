import {
	buildArtisticFilterMarkup,
	buildSimpleArtisticCss,
	ARTISTIC_PREVIEW_EFFECTS,
} from './picture-artistic-filters';
const NS = 'http://www.w3.org/2000/svg';
/** PowerPoint's live SVG/CSS artistic approximations; not a native Office picture cache. */
export function artisticPictureNodes(
	doc: Document,
	name: string,
	radius: number,
): SVGElement[] | undefined {
	if (!ARTISTIC_PREVIEW_EFFECTS.has(name) || !Number.isFinite(radius) || radius < 0 || radius > 100)
		return undefined;
	const markup = buildArtisticFilterMarkup(name, radius);
	if (markup) {
		const holder = doc.createElementNS(NS, 'svg');
		holder.innerHTML = markup;
		return Array.from(holder.children) as SVGElement[];
	}
	const css = buildSimpleArtisticCss(name, radius);
	if (!css) return undefined;
	const nodes: SVGElement[] = [];
	for (const match of css.matchAll(/([a-z-]+)\(([-\d.]+)(%|px|deg)?\)/g)) {
		const effect = match[1],
			value = Number(match[2]) / (match[3] === '%' ? 100 : 1);
		let node: SVGElement;
		if (effect === 'blur') {
			node = doc.createElementNS(NS, 'feGaussianBlur');
			node.setAttribute('stdDeviation', String(value));
		} else if (effect === 'saturate' || effect === 'grayscale') {
			node = doc.createElementNS(NS, 'feColorMatrix');
			node.setAttribute('type', 'saturate');
			node.setAttribute('values', String(effect === 'grayscale' ? 1 - value : value));
		} else if (effect === 'brightness' || effect === 'contrast') {
			node = doc.createElementNS(NS, 'feComponentTransfer');
			for (const ch of ['R', 'G', 'B']) {
				const fn = doc.createElementNS(NS, `feFunc${ch}`);
				fn.setAttribute('type', 'linear');
				fn.setAttribute('slope', String(value));
				fn.setAttribute('intercept', String(effect === 'contrast' ? (1 - value) / 2 : 0));
				node.append(fn);
			}
		} else if (effect === 'sepia') {
			node = doc.createElementNS(NS, 'feColorMatrix');
			node.setAttribute('type', 'matrix');
			node.setAttribute(
				'values',
				'0.393 0.769 0.189 0 0 0.349 0.686 0.168 0 0 0.272 0.534 0.131 0 0 0 0 0 1 0',
			);
		} else return undefined;
		nodes.push(node);
	}
	return nodes.length ? nodes : undefined;
}
