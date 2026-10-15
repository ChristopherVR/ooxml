// The horizontal and vertical rulers of Page Layout view: inches measured from the margin, with
// the margins shaded, following the first column and the rows of pages.
import { PAGE_VIEW_GAP, type PageViewLayout } from 'ooxml-core/xlsx';

const EIGHTH = 96 / 8;

function rule(
	doc: Document,
	host: HTMLElement,
	axis: 'h' | 'v',
	origin: number,
	margin: number,
	content: number,
	paper: number,
	z: number,
): void {
	const area = doc.createElement('div');
	area.className = 'xpl-ruler-area';
	const lead = origin + margin * z;
	if (axis === 'h') area.style.cssText = `left:${lead}px;width:${content * z}px;top:0;bottom:0`;
	else area.style.cssText = `top:${lead}px;height:${content * z}px;left:0;right:0`;
	host.append(area);
	const eighths = Math.floor(paper / EIGHTH);
	for (let i = 0; i <= eighths; i++) {
		const at = i * EIGHTH;
		const offset = at - margin;
		const whole =
			Math.abs(offset) < 0.01 || Math.abs(offset / 96 - Math.round(offset / 96)) < 0.001;
		const half = !whole && Math.abs(offset / 48 - Math.round(offset / 48)) < 0.001;
		const tick = doc.createElement('div');
		tick.className = 'xpl-tick';
		const size = whole ? 9 : half ? 6 : 3;
		const pos = origin + at * z;
		tick.style.cssText =
			axis === 'h' ? `left:${pos}px;height:${size}px` : `top:${pos}px;width:${size}px`;
		host.append(tick);
		if (whole && Math.abs(offset) > 0.01) {
			const label = doc.createElement('span');
			label.className = 'xpl-ruler-num';
			label.textContent = String(Math.abs(Math.round(offset / 96)));
			label.style.cssText = axis === 'h' ? `left:${pos + 2}px;top:2px` : `top:${pos + 2}px`;
			host.append(label);
		}
	}
}

/** Redraws both rulers for `layout` at zoom `z`; `pad` is the canvas padding around the pages. */
export function paintRulers(
	doc: Document,
	horizontal: HTMLElement,
	vertical: HTMLElement,
	layout: PageViewLayout,
	z: number,
	pad: number,
): void {
	horizontal.replaceChildren();
	vertical.replaceChildren();
	const { margins, paperWidth, paperHeight, contentWidth, contentHeight } = layout.plan;
	const cols = Math.max(...layout.pages.map((p) => p.slotCol)) + 1;
	const rows = Math.min(50, Math.max(...layout.pages.map((p) => p.slotRow)) + 1);
	const gap = PAGE_VIEW_GAP;
	for (let c = 0; c < Math.min(cols, 20); c++)
		rule(
			doc,
			horizontal,
			'h',
			(pad + c * (paperWidth + gap)) * z,
			margins.left * 96,
			contentWidth,
			paperWidth,
			z,
		);
	for (let r = 0; r < rows; r++)
		rule(
			doc,
			vertical,
			'v',
			(pad + r * (paperHeight + gap)) * z,
			margins.top * 96,
			contentHeight,
			paperHeight,
			z,
		);
}
