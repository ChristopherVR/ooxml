import type { VisioStyle } from './model';
import { number, type Cells } from './sheet';

/** Cached source paint state stays separate from gradient/tile render multipliers. */
export function sourcePaint(
	cells: Cells,
	resolveColor: (name: string, fallback: string) => string,
): Partial<VisioStyle> {
	const opacity = (name: string) => {
		if (cells.get(name)?.error) return undefined;
		const cached = number(cells, name, cells.has(name) ? NaN : 0);
		return Number.isFinite(cached) && cached >= 0 && cached <= 1 ? 1 - cached : undefined;
	};
	const foregroundOpacity = opacity('FillForegndTrans'),
		backgroundOpacity = opacity('FillBkgndTrans'),
		lineColorOpacity = opacity('LineColorTrans');
	const pattern = cells.get('FillPattern')?.error
		? NaN
		: number(cells, 'FillPattern', cells.has('FillPattern') ? NaN : 1);
	const backgroundColor = cells.get('FillBkgnd')?.error
		? ''
		: resolveColor('FillBkgnd', cells.has('FillBkgnd') ? '' : '#ffffff');
	return {
		...(Number.isInteger(pattern) && pattern >= 0 && pattern <= 65535
			? { fillPatternIndex: pattern }
			: {}),
		...(foregroundOpacity === undefined ? {} : { fillForegroundOpacity: foregroundOpacity }),
		...(backgroundColor ? { fillBackgroundColor: backgroundColor } : {}),
		...(backgroundOpacity === undefined ? {} : { fillBackgroundOpacity: backgroundOpacity }),
		...(lineColorOpacity === undefined ? {} : { lineColorOpacity }),
	};
}
