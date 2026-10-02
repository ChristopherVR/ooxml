import { elements } from '../xml/index.js';
import type { VisioLinearGradient } from './model.js';
import { number, sectionRows, type Report, type Sheet } from './sheet.js';
import { themePaintContext, type ThemeResources } from './theme-resolve.js';
import {
	colorChoice,
	drawingPaint,
	DRAWING_NS,
	integer,
	themeChild,
	themeChildren,
} from './theme-color.js';

/**
 * DrawingML ang is clockwise; Visio local coordinates are y-up. Project the
 * rectangle onto the unscaled direction to obtain the complete color range.
 * Visio ignores a:lin@scaled, a:gradFill@flip, and a:tileRect (MS-VSDX 2.3.4.2.22).
 * https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/5c862363-87dc-41d4-90d0-a892bf545fdf
 * https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.lineargradientfill
 * https://learn.microsoft.com/en-us/answers/questions/2248059/non-preset-(a-tilerect)-behaves-strange-in-case-of
 */
export function linearGradientEndpoints(
	width: number,
	height: number,
	angle: number,
): Pick<VisioLinearGradient, 'start' | 'end'> {
	const radians = ((angle / 60_000) * Math.PI) / 180;
	const clean = (value: number) => (Math.abs(value) < 1e-12 ? 0 : value);
	const dx = clean(Math.cos(radians)),
		dy = clean(-Math.sin(radians));
	const half = (Math.abs(dx) * width + Math.abs(dy) * height) / 2;
	return {
		start: [width / 2 - dx * half, height / 2 - dy * half],
		end: [width / 2 + dx * half, height / 2 + dy * half],
	};
}
/** Only complete, unscaled, shape-rotating linear themes are normalized. */
export function themeLinearGradient(
	sheet: Sheet,
	resources: ThemeResources,
	width: number,
	height: number,
	report: Report,
): VisioLinearGradient | undefined {
	const cells = sheet.cells;
	if (cells.get('FillForegnd')?.value !== 'Themed') return undefined;
	const enabled = cells.get('FillGradientEnabled');
	if (enabled && enabled.value !== 'Themed' && number(cells, 'FillGradientEnabled', 0) !== 1)
		return undefined;
	const context = themePaintContext(cells, 'Fill', resources, report);
	if (!context || context.selected.localName !== 'gradFill') return undefined;
	const reject = () => {
		report(
			'unsupported-theme-gradient',
			'The theme gradient has unsupported settings or overrides; its base color is used.',
		);
		return undefined;
	};
	if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0)
		return reject();
	if (
		['FillGradientDir', 'FillGradientAngle', 'RotateGradientWithShape', 'UseGroupGradient'].some(
			(key) => cells.has(key) && cells.get(key)?.value !== 'Themed',
		)
	)
		return reject();
	if (
		sectionRows(sheet, 'FillGradient').some((row) =>
			[...row.cells.values()].some((cell) => cell.value !== 'Themed'),
		)
	)
		return reject();
	const fill = context.selected;
	const lin = themeChild(fill, 'lin');
	const rawAngle = lin?.getAttribute('ang') ?? '0';
	const angle = /^\d{1,8}$/.test(rawAngle) ? Number(rawAngle) : NaN;
	if (
		!lin ||
		!Number.isInteger(angle) ||
		angle < 0 ||
		angle >= 21_600_000 ||
		!['1', 'true'].includes(fill.getAttribute('rotWithShape') ?? '1') ||
		elements(fill).some(
			(node) =>
				node.namespaceURI !== DRAWING_NS || !['gsLst', 'lin', 'tileRect'].includes(node.localName),
		) ||
		themeChildren(fill, 'lin').length !== 1 ||
		themeChildren(fill, 'gsLst').length !== 1
	)
		return reject();
	const nodes = themeChildren(themeChild(fill, 'gsLst'), 'gs');
	if (
		nodes.length < 2 ||
		nodes.length > 32 ||
		elements(themeChild(fill, 'gsLst')!).length !== nodes.length
	)
		return reject();
	const stops: VisioLinearGradient['stops'] = [];
	for (const node of nodes) {
		const position = integer(node.getAttribute('pos'));
		const paint = drawingPaint(colorChoice(node), context.colors, context.base);
		if (
			position === undefined ||
			position > 100_000 ||
			!paint ||
			position / 100_000 < (stops.at(-1)?.offset ?? 0)
		)
			return reject();
		stops.push({ offset: position / 100_000, ...paint });
	}
	return { type: 'linear', ...linearGradientEndpoints(width, height, angle), stops };
}
