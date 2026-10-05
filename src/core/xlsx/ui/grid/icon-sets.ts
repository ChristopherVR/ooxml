// Small SVGs for conditional-format icon sets (index 0 is the lowest bucket, as the core reports).
const RED = '#d9372b';
const YELLOW = '#f2b600';
const GREEN = '#3f9b3f';
const GRAY = '#7f7f7f';
const BLACK = '#404040';

const svg = (body: string, size: number): string =>
	`<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="0 0 16 16" aria-hidden="true">${body}</svg>`;

/** Arrow rotated by `deg` (0 = up). */
const arrow = (color: string, deg: number): string =>
	`<g transform="rotate(${deg} 8 8)"><path d="M8 1.5 14 8h-3.5v6.5h-5V8H2z" fill="${color}"/></g>`;

const circle = (color: string, rim = false): string =>
	`<circle cx="8" cy="8" r="${rim ? 5.5 : 6.5}" fill="${color}"${rim ? ' stroke="#333" stroke-width="2.4"' : ''}/>`;

const flag = (color: string): string =>
	`<path d="M3 1.5h1.3v13H3z" fill="#555"/><path d="M4.3 2h9l-2.6 3.5 2.6 3.5h-9z" fill="${color}"/>`;

const symbol = (kind: 'x' | '!' | 'v', color: string, circled: boolean): string => {
	const mark =
		kind === 'x'
			? '<path d="M5.2 5.2l5.6 5.6M10.8 5.2l-5.6 5.6" stroke="#fff" stroke-width="2" stroke-linecap="round"/>'
			: kind === '!'
				? '<path d="M8 4v5" stroke="#fff" stroke-width="2.2" stroke-linecap="round"/><circle cx="8" cy="11.8" r="1.2" fill="#fff"/>'
				: '<path d="M4.6 8.3l2.3 2.3 4.6-4.9" stroke="#fff" stroke-width="2" fill="none" stroke-linecap="round"/>';
	if (circled) return `<circle cx="8" cy="8" r="7" fill="${color}"/>${mark}`;
	return mark.replace(/#fff/g, color);
};

const sign = (index: number): string =>
	index === 0
		? `<path d="M8 1.5 14.5 8 8 14.5 1.5 8z" fill="${RED}"/>`
		: index === 1
			? `<path d="M8 2 15 14H1z" fill="${YELLOW}"/>`
			: circle(GREEN);

const star = (fill: 0 | 0.5 | 1): string => {
	const d = 'M8 1.2l2 4.3 4.7.6-3.4 3.2.9 4.7L8 11.7 3.8 14l.9-4.7L1.3 6.1 6 5.5z';
	const id = `s${fill === 0.5 ? 'h' : fill}`;
	return fill === 0.5
		? `<defs><linearGradient id="${id}"><stop offset=".5" stop-color="${YELLOW}"/><stop offset=".5" stop-color="#fff"/></linearGradient></defs><path d="${d}" fill="url(#${id})" stroke="${YELLOW}"/>`
		: `<path d="${d}" fill="${fill ? YELLOW : '#fff'}" stroke="${YELLOW}"/>`;
};

const bars = (filled: number, total: number): string => {
	let out = '';
	for (let i = 0; i < total; i++) {
		const h = 3 + (i * 10) / Math.max(1, total - 1);
		out += `<rect x="${1.5 + i * (13 / total)}" y="${14.5 - h}" width="${13 / total - 1}" height="${h}" fill="${i < filled ? '#2f6db5' : '#c8c8c8'}"/>`;
	}
	return out;
};

const quarters = (n: number): string => {
	if (n === 0) return '<circle cx="8" cy="8" r="6.5" fill="#fff" stroke="#555"/>';
	if (n === 4) return '<circle cx="8" cy="8" r="6.5" fill="#555"/>';
	const angle = (n / 4) * 2 * Math.PI;
	const x = 8 + 6.5 * Math.sin(angle);
	const y = 8 - 6.5 * Math.cos(angle);
	return `<circle cx="8" cy="8" r="6.5" fill="#fff" stroke="#555"/><path d="M8 8V1.5A6.5 6.5 0 0 1 ${x.toFixed(2)} ${y.toFixed(2)}z" fill="#555"/>`;
};

const boxes = (n: number): string => {
	let out = '';
	for (let i = 0; i < 4; i++) {
		const x = i % 2 ? 8.5 : 1.5;
		const y = i < 2 ? 1.5 : 8.5;
		out += `<rect x="${x}" y="${y}" width="6" height="6" fill="${i < n ? '#2f6db5' : '#d6d6d6'}"/>`;
	}
	return out;
};

function body(set: string, index: number): string {
	const arrows3 = [arrow(RED, 180), arrow(YELLOW, 90), arrow(GREEN, 0)];
	switch (set) {
		case '3Arrows':
			return arrows3[index] ?? '';
		case '3ArrowsGray':
			return arrow(GRAY, [180, 90, 0][index] ?? 0);
		case '3Flags':
			return flag([RED, YELLOW, GREEN][index] ?? GREEN);
		case '3TrafficLights2':
			return circle([RED, YELLOW, GREEN][index] ?? GREEN, true);
		case '3Signs':
			return sign(index);
		case '3Symbols':
			return symbol(
				(['x', '!', 'v'] as const)[index] ?? 'v',
				[RED, YELLOW, GREEN][index] ?? GREEN,
				true,
			);
		case '3Symbols2':
			return symbol(
				(['x', '!', 'v'] as const)[index] ?? 'v',
				[RED, YELLOW, GREEN][index] ?? GREEN,
				false,
			);
		case '3Stars':
			return star(([0, 0.5, 1] as const)[index] ?? 1);
		case '3Triangles':
			return index === 1
				? `<rect x="3" y="7" width="10" height="2.5" fill="${YELLOW}"/>`
				: `<path d="${index === 0 ? 'M2 5h12L8 13z' : 'M2 12h12L8 4z'}" fill="${index === 0 ? RED : GREEN}"/>`;
		case '4Arrows':
			return arrow([RED, YELLOW, YELLOW, GREEN][index] ?? GREEN, [180, 135, 45, 0][index] ?? 0);
		case '4ArrowsGray':
			return arrow(GRAY, [180, 135, 45, 0][index] ?? 0);
		case '4RedToBlack':
			return circle([BLACK, GRAY, '#f4a6a6', RED][index] ?? RED);
		case '4Rating':
			return bars(index + 1, 4);
		case '4TrafficLights':
			return circle([BLACK, RED, YELLOW, GREEN][index] ?? GREEN, true);
		case '5Arrows':
			return arrow(
				[RED, YELLOW, YELLOW, YELLOW, GREEN][index] ?? GREEN,
				[180, 135, 90, 45, 0][index] ?? 0,
			);
		case '5ArrowsGray':
			return arrow(GRAY, [180, 135, 90, 45, 0][index] ?? 0);
		case '5Rating':
			return bars(index, 4);
		case '5Quarters':
			return quarters(index);
		case '5Boxes':
			return boxes(index);
		default:
			return circle([RED, YELLOW, GREEN][Math.min(2, index)] ?? GREEN);
	}
}

/** The SVG markup of one icon of a set. */
export function iconSvg(set: string, index: number, size = 16): string {
	return svg(body(set, index), size);
}
