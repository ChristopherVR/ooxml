import type { InkCoordinate, SlideInkStroke } from 'ooxml-core/pptx/editor/ink';

const SVG_NS = 'http://www.w3.org/2000/svg';

export function drawOverlay(stage: HTMLElement): SVGSVGElement {
	const svg = stage.ownerDocument.createElementNS(SVG_NS, 'svg');
	svg.setAttribute('aria-hidden', 'true');
	svg.setAttribute('viewBox', `0 0 ${stage.offsetWidth} ${stage.offsetHeight}`);
	Object.assign(svg.style, {
		position: 'absolute',
		inset: '0',
		width: '100%',
		height: '100%',
		pointerEvents: 'none',
		overflow: 'hidden',
		zIndex: '20',
	});
	stage.append(svg);
	return svg;
}

export function inkOverlayPath(
	svg: SVGSVGElement,
	points: readonly InkCoordinate[],
): SVGPathElement {
	const path = svg.ownerDocument.createElementNS(SVG_NS, 'path');
	path.setAttribute(
		'd',
		points.map(({ x, y }, index) => `${index ? 'L' : 'M'}${x} ${y}`).join(' '),
	);
	path.setAttribute('fill', 'none');
	path.setAttribute('stroke-linejoin', 'round');
	path.setAttribute('stroke-linecap', 'round');
	svg.append(path);
	return path;
}

export function paintLasso(
	svg: SVGSVGElement,
	points: readonly InkCoordinate[],
	selected: readonly SlideInkStroke[],
) {
	svg.replaceChildren();
	for (const stroke of selected) {
		const path = inkOverlayPath(svg, stroke.points);
		path.setAttribute('stroke', 'var(--office-color-focus-ring, #0078d4)');
		path.setAttribute('stroke-width', String(Math.min(64, stroke.style.width + 5)));
		path.setAttribute('stroke-opacity', '0.55');
	}
	if (points.length) {
		const path = inkOverlayPath(svg, points);
		path.setAttribute('stroke', 'var(--office-color-focus-ring, #0078d4)');
		path.setAttribute('stroke-width', '2');
		path.setAttribute('stroke-dasharray', '5 4');
	}
}

export function stagePoint(stage: HTMLElement, event: PointerEvent): InkCoordinate {
	const rect = stage.getBoundingClientRect();
	return {
		x: ((event.clientX - rect.left) * stage.offsetWidth) / rect.width,
		y: ((event.clientY - rect.top) * stage.offsetHeight) / rect.height,
	};
}
