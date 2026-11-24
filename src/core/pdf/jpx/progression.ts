/**
 * Packet order (ITU-T T.800 B.12): the five progression orders (LRCP, RLCP, RPCL, PCRL, CPRL)
 * and progression order changes (POC). Each precinct's packets come in layer order, so a packet
 * is visited only when it is the next layer of its precinct; a POC volume that names packets
 * already read skips them.
 */
import type { Budget } from './errors';
import type { ProgressionChange } from './markers';
import type { Tile } from './structure';

export type VisitPacket = (layer: number, c: number, r: number, k: number) => void;

const divisible = (value: number, by: number) => (by > 2 ** 40 ? value === 0 : value % by === 0);

/** Calls `visit` for every packet of the tile in codestream order. */
export function progress(tile: Tile, budget: Budget, visit: VisitPacket): void {
	const counts = tile.components.map((component) =>
		component.resolutions.map((res) => res.precinctsWide * res.precinctsHigh),
	);
	const next = counts.map((perResolution) => perResolution.map((count) => new Int32Array(count)));
	const emit = (layer: number, c: number, r: number, k: number) => {
		const layers = next[c]![r]!;
		if (layers[k] !== layer) return;
		layers[k] = layer + 1;
		visit(layer, c, r, k);
	};
	const maxLevels = Math.max(...tile.components.map((c) => c.style.levels));
	const whole: ProgressionChange = {
		resolutionStart: 0,
		componentStart: 0,
		layerEnd: tile.style.layers,
		resolutionEnd: maxLevels + 1,
		componentEnd: tile.components.length,
		order: tile.style.order,
	};
	for (const change of [...tile.progression, whole]) {
		const volume = {
			...change,
			layerEnd: Math.min(change.layerEnd, tile.style.layers),
			resolutionEnd: Math.min(change.resolutionEnd, maxLevels + 1),
			componentEnd: Math.min(change.componentEnd, tile.components.length),
		};
		if (change.order > 4) continue;
		run(tile, volume, counts, budget, emit);
	}
}

function run(
	tile: Tile,
	v: ProgressionChange,
	counts: number[][],
	budget: Budget,
	emit: VisitPacket,
): void {
	const precincts = (c: number, r: number) => counts[c]?.[r] ?? 0;
	const components: number[] = [];
	for (let c = v.componentStart; c < v.componentEnd; c++) components.push(c);
	const resolutions: number[] = [];
	for (let r = v.resolutionStart; r < v.resolutionEnd; r++) resolutions.push(r);
	if (v.order === 0 || v.order === 1) {
		const layerFirst = v.order === 0;
		const outer = layerFirst ? v.layerEnd : resolutions.length;
		const inner = layerFirst ? resolutions.length : v.layerEnd;
		for (let a = 0; a < outer; a++)
			for (let b = 0; b < inner; b++) {
				const layer = layerFirst ? a : b;
				const r = resolutions[layerFirst ? b : a]!;
				for (const c of components) {
					const count = precincts(c, r);
					budget.spend(count);
					for (let k = 0; k < count; k++) emit(layer, c, r, k);
				}
			}
		return;
	}
	positional(tile, v, components, resolutions, budget, emit);
}

/** RPCL, PCRL and CPRL: precincts visited by their position on the reference grid (B.12.1.3). */
function positional(
	tile: Tile,
	v: ProgressionChange,
	components: number[],
	resolutions: number[],
	budget: Budget,
	emit: VisitPacket,
): void {
	let xStep = Infinity;
	let yStep = Infinity;
	const info = (c: number) => tile.components[c]!;
	for (const c of components)
		for (const r of resolutions) {
			const component = info(c);
			const res = component.resolutions[r];
			if (!res) continue;
			const levels = component.style.levels - r;
			xStep = Math.min(xStep, component.dx * 2 ** (res.ppx + levels));
			yStep = Math.min(yStep, component.dy * 2 ** (res.ppy + levels));
		}
	if (!Number.isFinite(xStep) || !Number.isFinite(yStep)) return;
	const layers = (c: number, r: number, x: number, y: number) => {
		const component = info(c);
		const res = component.resolutions[r];
		if (!res || !res.precinctsWide || !res.precinctsHigh) return;
		const levels = component.style.levels - r;
		const { dx, dy } = component;
		const scaleX = dx * 2 ** levels;
		const scaleY = dy * 2 ** levels;
		const rowOk =
			divisible(y, scaleY * 2 ** res.ppy) ||
			(y === tile.y0 && !divisible(res.y0 * 2 ** levels, 2 ** (res.ppy + levels)));
		const columnOk =
			divisible(x, scaleX * 2 ** res.ppx) ||
			(x === tile.x0 && !divisible(res.x0 * 2 ** levels, 2 ** (res.ppx + levels)));
		if (!rowOk || !columnOk) return;
		const px = Math.floor(Math.ceil(x / scaleX) / 2 ** res.ppx) - Math.floor(res.x0 / 2 ** res.ppx);
		const py = Math.floor(Math.ceil(y / scaleY) / 2 ** res.ppy) - Math.floor(res.y0 / 2 ** res.ppy);
		if (px < 0 || py < 0 || px >= res.precinctsWide || py >= res.precinctsHigh) return;
		const k = py * res.precinctsWide + px;
		for (let layer = 0; layer < v.layerEnd; layer++) emit(layer, c, r, k);
	};
	const positions = (each: (x: number, y: number) => void) => {
		for (let y = tile.y0; y < tile.y1; y += yStep - (y % yStep))
			for (let x = tile.x0; x < tile.x1; x += xStep - (x % xStep)) {
				budget.spend(1);
				each(x, y);
			}
	};
	if (v.order === 2)
		for (const r of resolutions)
			positions((x, y) => {
				for (const c of components) layers(c, r, x, y);
			});
	else if (v.order === 3)
		positions((x, y) => {
			for (const c of components) for (const r of resolutions) layers(c, r, x, y);
		});
	else
		for (const c of components)
			positions((x, y) => {
				for (const r of resolutions) layers(c, r, x, y);
			});
}
