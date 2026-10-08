// Print Layout pictures: package media through `PictureUrl`, charts painted by the core
// (`LayoutObject.chart`), and a placeholder box for anything that cannot be shown.

/** Resolves a package picture part to a displayable URL; undefined draws a placeholder box. */
export type PictureUrl = (partName: string, contentType: string) => string | undefined;

/** A chart painted by the core (`LayoutObject.chart`), or its labelled frame. */
function chartElement(chart: { svg?: string; label: string }): HTMLElement {
	const element = document.createElement('div');
	element.dataset.docxChart = chart.svg ? '1' : 'frame';
	const svg = chart.svg
		? new DOMParser().parseFromString(chart.svg, 'image/svg+xml').documentElement
		: undefined;
	if (svg && svg.localName === 'svg') {
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
	},
	pictureUrl: PictureUrl | undefined,
): HTMLElement {
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
