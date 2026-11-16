/**
 * Gallery tiles for Office's built-in SmartArt colour and quick-style definitions (the full
 * Change Colors and SmartArt Styles catalogue of `ooxml-core/diagram`): three boxes coloured the
 * way a definition's `node1` label fills and outlines them, with its effect strength, bevel and
 * spatial camera suggested. A tile is a hint for picking, not a rendering of the definition.
 */
import type { DiagramNativeColorEntry, DiagramNativeStyleEntry } from 'ooxml-core/diagram';
import { safeColor, svgTile } from '../ribbon/svg-preview';
import type { SmartArtTileSize } from './tiles';

/** Resolved theme colours by scheme name (`accent1`, `lt1`, `dk2`, ...), as CSS colours. */
export type SmartArtThemeColors = Partial<Record<string, string>>;

const DEFAULTS: Readonly<Record<string, string>> = {
	dk1: '#000000',
	lt1: '#FFFFFF',
	dk2: '#44546A',
	lt2: '#E7E6E6',
	accent1: '#4472C4',
	accent2: '#ED7D31',
	accent3: '#A5A5A5',
	accent4: '#FFC000',
	accent5: '#5B9BD5',
	accent6: '#70AD47',
};

const round = (value: number): string => String(Math.round(value * 100) / 100);

function rgb(color: string): [number, number, number] {
	const hex = color.replace('#', '');
	const full = hex.length === 3 ? [...hex].map((c) => c + c).join('') : hex.slice(0, 6);
	const value = Number.parseInt(full, 16);
	return Number.isNaN(value) ? [68, 114, 196] : [value >> 16, (value >> 8) & 255, value & 255];
}

const hex = (channels: readonly number[]): string =>
	`#${channels
		.map((c) =>
			Math.round(Math.min(255, Math.max(0, c)))
				.toString(16)
				.padStart(2, '0'),
		)
		.join('')}`;

/** Mix toward black (negative tone) or white (positive). */
function toned(color: string, tone = 0): string {
	const channels = rgb(color);
	return hex(channels.map((c) => (tone < 0 ? c * (1 + tone) : c + (255 - c) * tone)));
}

function mix(a: string, b: string, t: number): string {
	const [x, y] = [rgb(a), rgb(b)];
	return hex(x.map((c, i) => c + ((y[i] ?? c) - c) * t));
}

const resolve = (theme: SmartArtThemeColors, name: string): string =>
	safeColor(theme[name] ?? DEFAULTS[name], DEFAULTS.accent1!);

/**
 * The fills a colour definition's `node1` label gives `count` nodes: `span` blends the list across
 * the nodes, `cycle` repeats it, `repeat` keeps the last colour. An approximation of DiagramML's
 * colour application (no hue direction), for tiles and for an editor that recolours shapes.
 */
export function smartArtDefinitionFills(
	entry: DiagramNativeColorEntry,
	theme: SmartArtThemeColors,
	count = 3,
): string[] {
	const { fills, tones, mode } = entry.preview;
	const colors = fills.map((name, i) => toned(resolve(theme, name), tones?.[i]));
	if (!colors.length) return [resolve(theme, 'accent1')];
	return Array.from({ length: Math.max(1, count) }, (_, i) => {
		if (mode === 'span' && colors.length > 1) {
			const at = (count > 1 ? i / (count - 1) : 0) * (colors.length - 1);
			const low = Math.floor(at);
			return mix(colors[low]!, colors[Math.min(low + 1, colors.length - 1)]!, at - low);
		}
		return mode === 'cycle' ? colors[i % colors.length]! : colors[Math.min(i, colors.length - 1)]!;
	});
}

interface TileLook {
	fills: readonly string[];
	line: string;
	lineWidth: number;
	opacity: number;
	effect: number;
	bevel: boolean;
	spatial: boolean;
}

function tile(id: string, look: TileLook, size: SmartArtTileSize): string {
	const safeId = `sad-${id}`.replace(/[^\w-]/gu, '_');
	const pad = 4;
	const gap = 4;
	const w = (size.width - pad * 2 - gap * 2) / 3;
	const h = size.height * 0.46;
	const y = (size.height - h) / 2;
	let defs = '';
	let filter = '';
	if (look.effect > 0) {
		const k = Math.min(look.effect, 3);
		defs += `<filter id="${safeId}-s" x="-30%" y="-30%" width="160%" height="170%"><feDropShadow dx="0" dy="${round(0.5 * k)}" stdDeviation="${round(0.4 * k)}" flood-color="#000000" flood-opacity="${round(0.15 + 0.1 * k)}"/></filter>`;
		filter = ` filter="url(#${safeId}-s)"`;
	}
	if (look.bevel)
		defs += `<linearGradient id="${safeId}-b" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#FFFFFF" stop-opacity="0.45"/><stop offset="0.5" stop-color="#FFFFFF" stop-opacity="0"/><stop offset="1" stop-color="#000000" stop-opacity="0.2"/></linearGradient>`;
	const boxes = [0, 1, 2].map((i) => {
		const x = pad + i * (w + gap);
		const fill = safeColor(look.fills[i % look.fills.length], DEFAULTS.accent1!);
		const attrs = `x="${round(x)}" y="${round(y)}" width="${round(w)}" height="${round(h)}" rx="${round(Math.min(w, h) * 0.12)}"`;
		const body = `<rect ${attrs} fill="${fill}" fill-opacity="${round(look.opacity)}" stroke="${safeColor(look.line, '#FFFFFF')}" stroke-width="${round(look.lineWidth)}"${filter}/>`;
		return look.bevel ? `${body}<rect ${attrs} fill="url(#${safeId}-b)"/>` : body;
	});
	const group = look.spatial
		? `<g transform="translate(0 ${round(size.height * 0.08)}) skewY(-6)">${boxes.join('')}</g>`
		: boxes.join('');
	const bg = `<rect x="0" y="0" width="${size.width}" height="${size.height}" fill="#FFFFFF"/>`;
	return svgTile(size.width, size.height, defs, bg + group);
}

/** A Change Colors tile for a built-in colour definition. */
export function smartArtColorDefinitionTileSvg(
	entry: DiagramNativeColorEntry,
	theme: SmartArtThemeColors,
	size: SmartArtTileSize,
): string {
	const line = resolve(theme, entry.preview.line);
	const outlined = entry.preview.line !== 'lt1';
	return tile(
		entry.id,
		{
			fills: smartArtDefinitionFills(entry, theme),
			line,
			lineWidth: outlined ? 1.5 : 0.75,
			opacity: entry.preview.transparent ? 0.6 : 1,
			effect: 0,
			bevel: false,
			spatial: false,
		},
		size,
	);
}

/** A SmartArt Styles tile for a built-in quick-style definition, drawn in `accent`. */
export function smartArtStyleDefinitionTileSvg(
	entry: DiagramNativeStyleEntry,
	accent: string | undefined,
	size: SmartArtTileSize,
): string {
	const { line, effect, bevel, camera } = entry.preview;
	return tile(
		entry.id,
		{
			fills: [safeColor(accent, DEFAULTS.accent1!)],
			line: '#FFFFFF',
			lineWidth: line >= 3 ? 1.75 : line > 0 ? 0.75 : 0,
			opacity: 1,
			effect,
			bevel: bevel === true,
			spatial: camera !== undefined,
		},
		size,
	);
}
