/**
 * Prepare an inline `<svg>` of a mounted slide for the shared SVG to PDF converter, which
 * reads presentation attributes: a detached copy whose painting properties come from the
 * computed style (CSS rules and inline styles included), whose hidden nodes are dropped, and
 * which carries copies of the gradients, markers and clip paths it references from elsewhere
 * on the page. Also the matrix that places its user units in its CSS box.
 */
const PROPERTIES = [
	'fill',
	'fill-opacity',
	'fill-rule',
	'stroke',
	'stroke-width',
	'stroke-opacity',
	'stroke-dasharray',
	'stroke-linecap',
	'stroke-linejoin',
	'opacity',
	'font-family',
	'font-size',
	'font-weight',
	'font-style',
	'text-anchor',
	'dominant-baseline',
	'text-decoration',
	'letter-spacing',
	'clip-path',
	'filter',
	'mask',
	'marker-start',
	'marker-end',
	'stop-color',
	'stop-opacity',
] as const;

/** `url("http://page#id")` and `url(#id)` as `url(#id)`. */
const localUrl = (value: string) =>
	value.replace(/url\(\s*['"]?[^#'")]*#([^'")]+)['"]?\s*\)/g, 'url(#$1)');

function copyStyles(live: Element, copy: Element, view: Window): void {
	const style = view.getComputedStyle(live);
	for (const name of PROPERTIES) {
		let value = style.getPropertyValue(name).trim();
		if (!value) continue;
		if (name === 'text-decoration') value = style.getPropertyValue('text-decoration-line');
		if (value === 'normal' && name === 'letter-spacing') continue;
		if ((name === 'clip-path' || name === 'filter' || name === 'mask') && value === 'none') {
			copy.removeAttribute(name);
			continue;
		}
		copy.setAttribute(name, localUrl(value));
	}
	if (style.display === 'none' || style.visibility === 'hidden')
		copy.setAttribute('display', 'none');
	copy.removeAttribute('style');
	copy.removeAttribute('class');
}

const REFERENCE = /url\(#([^)]+)\)/g;

/** A detached, attribute-styled copy of `svg` (see the module comment). */
export function normalizeInlineSvg(svg: SVGSVGElement): Element {
	const doc = svg.ownerDocument;
	const view = doc.defaultView!;
	const copy = svg.cloneNode(true) as SVGSVGElement;
	const live = [svg, ...Array.from(svg.querySelectorAll('*'))];
	const copies = [copy, ...Array.from(copy.querySelectorAll('*'))];
	live.forEach((node, index) => {
		const target = copies[index];
		if (target) copyStyles(node, target, view);
	});
	for (const node of Array.from(copy.querySelectorAll('[display="none"]'))) {
		// Definitions are referenced, not drawn: keep them even when hidden.
		if (
			!node.closest('defs, clipPath, linearGradient, radialGradient, marker, pattern, mask, symbol')
		)
			node.remove();
	}
	// Pull in referenced definitions that live outside this <svg>.
	const present = new Set(Array.from(copy.querySelectorAll('[id]')).map((node) => node.id));
	const wanted = new Set<string>();
	for (const node of [copy, ...Array.from(copy.querySelectorAll('*'))]) {
		for (const attr of Array.from(node.attributes)) {
			for (const match of attr.value.matchAll(REFERENCE)) wanted.add(match[1]!);
			if ((attr.localName === 'href' || attr.name === 'xlink:href') && attr.value.startsWith('#'))
				wanted.add(attr.value.slice(1));
		}
	}
	let defs: Element | undefined;
	for (const id of wanted) {
		if (present.has(id)) continue;
		const source = doc.getElementById(id);
		if (!source || svg.contains(source)) continue;
		defs ??= copy.insertBefore(
			doc.createElementNS('http://www.w3.org/2000/svg', 'defs'),
			copy.firstChild,
		);
		const borrowed = source.cloneNode(true) as Element;
		const sourceNodes = [source, ...Array.from(source.querySelectorAll('*'))];
		const borrowedNodes = [borrowed, ...Array.from(borrowed.querySelectorAll('*'))];
		sourceNodes.forEach((node, index) => {
			const target = borrowedNodes[index];
			if (target) copyStyles(node, target, view);
		});
		defs.appendChild(borrowed);
		present.add(id);
	}
	return copy;
}

export { viewBoxMatrix } from 'ooxml-core/pdf/write';
