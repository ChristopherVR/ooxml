/** Format-neutral CT_ManualLayout values. Raw extensions stay with the source XML model. */
export interface ChartManualLayout {
	layoutTarget?: 'inner' | 'outer';
	xMode?: 'edge' | 'factor';
	yMode?: 'edge' | 'factor';
	widthMode?: 'edge' | 'factor';
	heightMode?: 'edge' | 'factor';
	x?: number;
	y?: number;
	width?: number;
	height?: number;
}

/** One reader for DOM and object-tree adapters; reject malformed values without losing source XML. */
export function readChartManualLayoutValues(
	value: (name: string) => unknown,
): ChartManualLayout | undefined {
	const result: ChartManualLayout = {};
	const enumeration = <T extends string>(name: string, values: readonly T[]): T | undefined => {
		const raw = value(name);
		return values.includes(raw as T) ? (raw as T) : undefined;
	};
	const target = enumeration('layoutTarget', ['inner', 'outer']);
	if (target) result.layoutTarget = target;
	for (const [xml, key] of [
		['xMode', 'xMode'],
		['yMode', 'yMode'],
		['wMode', 'widthMode'],
		['hMode', 'heightMode'],
	] as const) {
		const mode = enumeration(xml, ['edge', 'factor']);
		if (mode) result[key] = mode;
	}
	for (const [xml, key] of [
		['x', 'x'],
		['y', 'y'],
		['w', 'width'],
		['h', 'height'],
	] as const) {
		const raw = value(xml);
		if (
			(typeof raw !== 'string' && typeof raw !== 'number') ||
			(typeof raw === 'string' && !raw.trim())
		)
			continue;
		const number = Number(raw);
		if (Number.isFinite(number)) result[key] = number;
	}
	return Object.keys(result).length ? result : undefined;
}

export interface ChartLayoutRect {
	x: number;
	y: number;
	width: number;
	height: number;
}
export interface ChartFrameSize {
	width: number;
	height: number;
}

export function hasManualLayoutFields(
	layout: ChartManualLayout | null | undefined,
): layout is ChartManualLayout {
	return (
		layout != null &&
		(layout.x !== undefined ||
			layout.y !== undefined ||
			layout.width !== undefined ||
			layout.height !== undefined)
	);
}

/** Fractions of the chart frame: edge is absolute, factor offsets an automatic position. */
export function resolveManualLayoutRect(
	layout: ChartManualLayout | null | undefined,
	frame: ChartFrameSize,
	auto: ChartLayoutRect,
): ChartLayoutRect | undefined {
	if (!hasManualLayoutFields(layout)) return undefined;
	const x =
		layout.x === undefined
			? auto.x
			: layout.xMode === 'edge'
				? layout.x * frame.width
				: auto.x + layout.x * frame.width;
	const y =
		layout.y === undefined
			? auto.y
			: layout.yMode === 'edge'
				? layout.y * frame.height
				: auto.y + layout.y * frame.height;
	const width =
		layout.width === undefined
			? auto.width
			: layout.widthMode === 'edge'
				? layout.width * frame.width - x
				: layout.width * frame.width;
	const height =
		layout.height === undefined
			? auto.height
			: layout.heightMode === 'edge'
				? layout.height * frame.height - y
				: layout.height * frame.height;
	return { x, y, width: Math.max(width, 1), height: Math.max(height, 1) };
}
