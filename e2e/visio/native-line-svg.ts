import type { Page } from '@playwright/test';

/** Independent physical-page endpoints from native Visio's rounded SVG export. */
export async function nativeSvgLineEndpoints(page: Page, source: string) {
	return page.evaluate((source) => {
		const parsed = new DOMParser().parseFromString(source, 'image/svg+xml');
		if (parsed.querySelector('parsererror')) throw new Error('Invalid native SVG capture.');
		const svg = document.importNode(parsed.documentElement, true) as unknown as SVGSVGElement;
		svg.style.position = 'absolute';
		svg.style.left = '-10000px';
		svg.style.visibility = 'hidden';
		document.body.append(svg);
		try {
			const dimension = (name: string) => {
				const match = /^([\d.]+)in$/.exec(svg.getAttribute(name) ?? '');
				if (!match) throw new Error('Native SVG requires physical inch dimensions.');
				return Number(match[1]);
			};
			const width = dimension('width'),
				height = dimension('height');
			const box = svg.viewBox.baseVal;
			const root = svg.getScreenCTM()!.inverse();
			const lines = Array.from(svg.querySelectorAll<SVGGElement>('g[id]')).map((group) => {
				const id = /^shape(\d+)-/.exec(group.id)?.[1];
				const path = group.querySelector<SVGGeometryElement>(
					'path,rect,polygon,line,ellipse,circle,polyline',
				);
				if (!id || !path) throw new Error('Unexpected native shape SVG group.');
				const matrix = root.multiply(path.getScreenCTM()!);
				const point = (length: number) => {
					const local = path.getPointAtLength(length);
					const transformed = new DOMPoint(local.x, local.y).matrixTransform(matrix);
					return {
						x: ((transformed.x - box.x) * width) / box.width,
						y: height - ((transformed.y - box.y) * height) / box.height,
					};
				};
				const style = getComputedStyle(path);
				return {
					id,
					begin: point(0),
					end: point(path.getTotalLength()),
					paint: {
						fill: style.fill,
						stroke: style.stroke,
						fillOpacity: style.fillOpacity,
						strokeOpacity: style.strokeOpacity,
					},
				};
			});
			return { width, height, lines };
		} finally {
			svg.remove();
		}
	}, source);
}
