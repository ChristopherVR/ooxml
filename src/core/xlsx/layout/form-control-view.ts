// A static appearance for each legacy form control, as SVG (the Windows look Excel gives them).
// The editor paints it and maps clicks to `applyControlAction` through `data-control-*` attributes
// on the parts that react: `toggle`, `item` (with an index), `step-up` (increase), `step-down`,
// `track` (a page step toward the click), `thumb` and `drop`.
import type { FormControlObject } from '../model';
import { controlBounds, type ControlState } from './form-control-state';

export interface FormControlViewOptions {
	/** Items of a drop-down or list box, from its list range. */
	items?: readonly string[];
	/** Pixels per pixel at 100%: the sheet zoom / 100. */
	scale?: number;
}

const FACE = '#e1e1e1';
const EDGE = '#adadad';
const TEXT = '#000000';
const FONT = 'font-family="Segoe UI,Arial,sans-serif"';
const round = (value: number) => Math.round(value * 100) / 100;
const escapeXml = (text: string) =>
	text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

function caption(text: string | undefined, x: number, h: number, size: number) {
	if (!text) return '';
	return `<text x="${round(x)}" y="${round(h / 2)}" ${FONT} font-size="${round(size)}" fill="${TEXT}" dominant-baseline="central">${escapeXml(text)}</text>`;
}

function checkbox(c: FormControlObject, s: ControlState, w: number, h: number, size: number) {
	const box = Math.min(13 * (size / 10.67), h - 2);
	const y = (h - box) / 2;
	const mark =
		s.checked === 'checked'
			? `<path d="M${round(2 + box * 0.2)} ${round(y + box * 0.5)} l${round(box * 0.25)} ${round(box * 0.25)} l${round(box * 0.5)} ${round(-box * 0.5)}" fill="none" stroke="${TEXT}" stroke-width="1.6"/>`
			: s.checked === 'mixed'
				? `<rect x="${round(2 + box * 0.25)}" y="${round(y + box * 0.25)}" width="${round(box * 0.5)}" height="${round(box * 0.5)}" fill="#808080"/>`
				: '';
	return `<g data-control-part="toggle"><rect x="2" y="${round(y)}" width="${round(box)}" height="${round(box)}" fill="#fff" stroke="#333"/>${mark}${caption(c.caption, box + 8, h, size)}</g>`;
}

function radio(c: FormControlObject, s: ControlState, w: number, h: number, size: number) {
	const d = Math.min(13 * (size / 10.67), h - 2);
	const dot =
		s.checked === 'checked'
			? `<circle cx="${round(2 + d / 2)}" cy="${round(h / 2)}" r="${round(d * 0.25)}" fill="${TEXT}"/>`
			: '';
	return `<g data-control-part="toggle"><circle cx="${round(2 + d / 2)}" cy="${round(h / 2)}" r="${round(d / 2)}" fill="#fff" stroke="#333"/>${dot}${caption(c.caption, d + 8, h, size)}</g>`;
}

function arrow(cx: number, cy: number, up: boolean, size = 3) {
	const dy = up ? -size / 2 : size / 2;
	return `<path d="M${round(cx - size)} ${round(cy - dy)} L${round(cx)} ${round(cy + dy)} L${round(cx + size)} ${round(cy - dy)} Z" fill="#606060"/>`;
}

function dropdown(
	c: FormControlObject,
	s: ControlState,
	items: readonly string[],
	w: number,
	h: number,
	size: number,
) {
	const button = Math.min(16, w / 2);
	const text = s.selected >= 1 ? (items[s.selected - 1] ?? '') : '';
	const clip = `<clipPath id="fc-drop-${c.shapeId}"><rect x="2" y="1" width="${round(Math.max(0, w - button - 4))}" height="${round(h - 2)}"/></clipPath>`;
	return `<defs>${clip}</defs><rect x="0.5" y="0.5" width="${round(w - 1)}" height="${round(h - 1)}" fill="#fff" stroke="#7a7a7a"/><g clip-path="url(#fc-drop-${c.shapeId})">${caption(text, 4, h, size)}</g><rect data-control-part="drop" x="${round(w - button - 0.5)}" y="1.5" width="${round(button - 1)}" height="${round(h - 3)}" fill="${FACE}" stroke="${EDGE}"/>${arrow(w - button / 2, h / 2, false)}`;
}

function list(s: ControlState, items: readonly string[], w: number, h: number, size: number) {
	const row = size * 1.4;
	const rows = Math.max(0, Math.floor((h - 2) / row));
	const lines = items.slice(0, rows).map((item, index) => {
		const y = 1 + index * row;
		const selected = s.selected === index + 1;
		return `<g data-control-part="item" data-control-index="${index + 1}"><rect x="1" y="${round(y)}" width="${round(w - 2)}" height="${round(row)}" fill="${selected ? '#0078d7' : '#fff'}"/><text x="4" y="${round(y + row / 2)}" ${FONT} font-size="${round(size)}" fill="${selected ? '#fff' : TEXT}" dominant-baseline="central">${escapeXml(item)}</text></g>`;
	});
	return `<rect x="0.5" y="0.5" width="${round(w - 1)}" height="${round(h - 1)}" fill="#fff" stroke="#7a7a7a"/>${lines.join('')}`;
}

function button(c: FormControlObject, w: number, h: number, size: number) {
	return `<rect x="0.5" y="0.5" width="${round(w - 1)}" height="${round(h - 1)}" fill="${FACE}" stroke="${EDGE}"/><text x="${round(w / 2)}" y="${round(h / 2)}" ${FONT} font-size="${round(size)}" fill="${TEXT}" text-anchor="middle" dominant-baseline="central">${escapeXml(c.caption ?? '')}</text>`;
}

function spin(w: number, h: number) {
	const half = h / 2;
	return `<rect data-control-part="step-up" x="0.5" y="0.5" width="${round(w - 1)}" height="${round(half - 0.5)}" fill="${FACE}" stroke="${EDGE}"/>${arrow(w / 2, half / 2, true)}<rect data-control-part="step-down" x="0.5" y="${round(half)}" width="${round(w - 1)}" height="${round(half - 0.5)}" fill="${FACE}" stroke="${EDGE}"/>${arrow(w / 2, half + half / 2, false)}`;
}

function scroll(c: FormControlObject, s: ControlState, w: number, h: number) {
	const horizontal = c.horizontal ?? w >= h;
	const { min, max } = controlBounds(c);
	const ratio = max > min ? (s.value - min) / (max - min) : 0;
	const end = Math.min(16, horizontal ? h : w);
	const length = (horizontal ? w : h) - 2 * end;
	const thumbLength = Math.max(8, Math.min(length, end));
	const offset = end + ratio * Math.max(0, length - thumbLength);
	const track = '#f0f0f0';
	const rect = (part: string, x: number, y: number, rw: number, rh: number, fill: string) =>
		`<rect data-control-part="${part}" x="${round(x)}" y="${round(y)}" width="${round(rw)}" height="${round(rh)}" fill="${fill}" stroke="${EDGE}"/>`;
	if (horizontal)
		return `${rect('track', end, 0.5, length, h - 1, track)}${rect('step-down', 0.5, 0.5, end, h - 1, FACE)}${rect('step-up', w - end - 0.5, 0.5, end, h - 1, FACE)}${rect('thumb', offset, 0.5, thumbLength, h - 1, FACE)}<path d="M${round(end * 0.6)} ${round(h / 2 - 3)} l-3 3 l3 3 Z M${round(w - end * 0.6)} ${round(h / 2 - 3)} l3 3 l-3 3 Z" fill="#606060"/>`;
	return `${rect('track', 0.5, end, w - 1, length, track)}${rect('step-down', 0.5, 0.5, w - 1, end, FACE)}${rect('step-up', 0.5, h - end - 0.5, w - 1, end, FACE)}${rect('thumb', 0.5, offset, w - 1, thumbLength, FACE)}${arrow(w / 2, end / 2, true)}${arrow(w / 2, h - end / 2, false)}`;
}

function groupBox(c: FormControlObject, w: number, h: number, size: number) {
	const label = c.caption ?? '';
	const gap = label ? label.length * size * 0.55 + 6 : 0;
	const top = size / 2;
	return `<path d="M8 ${round(top)} H4 a2 2 0 0 0 -2 2 V${round(h - 3)} a2 2 0 0 0 2 2 H${round(w - 4)} a2 2 0 0 0 2 -2 V${round(top + 2)} a2 2 0 0 0 -2 -2 H${round(10 + gap)}" fill="none" stroke="#a0a0a0"/><text x="10" y="${round(top)}" ${FONT} font-size="${round(size)}" fill="${TEXT}" dominant-baseline="central">${escapeXml(label)}</text>`;
}

/** The control drawn into `widthPx` by `heightPx`, or undefined for a type with no appearance. */
export function formControlSvg(
	control: FormControlObject,
	state: ControlState,
	widthPx: number,
	heightPx: number,
	options: FormControlViewOptions = {},
): string | undefined {
	const w = Math.max(1, widthPx);
	const h = Math.max(1, heightPx);
	const size = 10.67 * (options.scale ?? 1);
	const items = options.items ?? [];
	let body: string;
	switch (control.controlType) {
		case 'checkbox':
			body = checkbox(control, state, w, h, size);
			break;
		case 'radio':
			body = radio(control, state, w, h, size);
			break;
		case 'dropdown':
			body = dropdown(control, state, items, w, h, size);
			break;
		case 'list':
			body = list(state, items, w, h, size);
			break;
		case 'button':
			body = button(control, w, h, size);
			break;
		case 'spin':
			body = spin(w, h);
			break;
		case 'scroll':
			body = scroll(control, state, w, h);
			break;
		case 'groupBox':
			body = groupBox(control, w, h, size);
			break;
		case 'label':
			body = caption(control.caption, 2, h, size);
			break;
		default:
			return undefined;
	}
	const label = escapeXml(control.caption ?? control.name ?? control.controlType);
	return `<svg xmlns="http://www.w3.org/2000/svg" width="${round(w)}" height="${round(h)}" viewBox="0 0 ${round(w)} ${round(h)}" role="img" aria-label="${label}">${body}</svg>`;
}
