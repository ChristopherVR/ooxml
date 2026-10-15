// Print Layout pictures: package media through `PictureUrl`, charts painted by the core
// (`LayoutObject.chart`), and a placeholder box for anything that cannot be shown.

/** Resolves a package picture part to a displayable URL; undefined draws a placeholder box. */
export type PictureUrl = (partName: string, contentType: string) => string | undefined;

/**
 * Whether an SVG `href` / `xlink:href` may stay: an in-document fragment or an
 * embedded raster image. Anything else (a script URL, a remote resource) goes.
 */
function isSafeSvgHref(value: string): boolean {
	const trimmed = value.trim();
	return trimmed.startsWith('#') || /^data:image\//iu.test(trimmed);
}

/**
 * Strips the active content from a chart SVG before it is imported into the
 * page: `<script>` and `<foreignObject>` elements, every `on*` event-handler
 * attribute, and `href` / `xlink:href` values other than fragments or
 * `data:image` URLs. Mirrors the rules of `src/core/docx/ui/chart-spec.ts`.
 */
export function sanitizeChartSvg(svg: Element): void {
	const elements = [svg, ...Array.from(svg.getElementsByTagName('*'))];
	for (const element of elements) {
		const name = element.localName.toLowerCase();
		if (name === 'script' || name === 'foreignobject') {
			element.remove();
			continue;
		}
		for (const attr of Array.from(element.attributes)) {
			const unsafeHref = attr.localName === 'href' && !isSafeSvgHref(attr.value);
			if (/^on/iu.test(attr.name) || unsafeHref) {
				element.removeAttribute(attr.name);
			}
		}
	}
}

/** A chart painted by the core (`LayoutObject.chart`), or its labelled frame. */
function chartElement(chart: { svg?: string; label: string }): HTMLElement {
	const element = document.createElement('div');
	element.dataset.docxChart = chart.svg ? '1' : 'frame';
	const svg = chart.svg
		? new DOMParser().parseFromString(chart.svg, 'image/svg+xml').documentElement
		: undefined;
	if (svg && svg.localName === 'svg') {
		sanitizeChartSvg(svg);
		element.classList.add('dve-print-chart');
		element.title = chart.label;
		element.append(document.importNode(svg, true));
	} else {
		element.classList.add('dve-print-picture-missing', 'dve-print-chart-frame');
		element.textContent = chart.label;
	}
	return element;
}

export function pictureElement(
	object: {
		partName: string;
		contentType: string;
		widthPx: number;
		heightPx: number;
		chart?: { svg?: string; label: string };
		noteSeparator?: 'separator' | 'continuationSeparator';
	},
	pictureUrl: PictureUrl | undefined,
): HTMLElement {
	if (object.noteSeparator) {
		const rule = document.createElement('span');
		rule.className = 'dve-print-note-rule';
		Object.assign(rule.style, {
			position: 'absolute',
			width: `${object.widthPx}px`,
			height: '1px',
			borderTop: '1px solid currentColor',
		});
		return rule;
	}
	if (object.chart) {
		const element = chartElement(object.chart);
		element.classList.add('dve-print-picture');
		element.style.position = 'absolute';
		element.style.width = `${object.widthPx}px`;
		element.style.height = `${object.heightPx}px`;
		return element;
	}
	const url = object.partName ? pictureUrl?.(object.partName, object.contentType) : undefined;
	const element = document.createElement(url ? 'img' : 'div');
	if (url) {
		(element as HTMLImageElement).src = url;
		(element as HTMLImageElement).alt = '';
	} else element.classList.add('dve-print-picture-missing');
	element.classList.add('dve-print-picture');
	element.style.position = 'absolute';
	element.style.width = `${object.widthPx}px`;
	element.style.height = `${object.heightPx}px`;
	return element;
}
