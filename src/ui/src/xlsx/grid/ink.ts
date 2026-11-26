// Ink over the cells: its strokes as SVG, stretched onto the anchor's box as Excel draws them.
import { inkObjectSvg, type InkObject } from 'ooxml-core/xlsx';
import type { Box } from 'ooxml-core/xlsx/ui';
import { svgNode } from './dom';
import type { GridView } from './grid-view';

export function paintInk(view: GridView, node: HTMLElement, ink: InkObject, box: Box): void {
	node.dataset.ink = '';
	node.classList.add('xg-ink');
	node.setAttribute('role', 'img');
	node.setAttribute('aria-label', ink.description || ink.name || view.ctx.t('Ink'));
	const svg = svgNode(view.doc, inkObjectSvg(ink, box.w, box.h));
	svg.setAttribute('aria-hidden', 'true');
	node.append(svg);
}
