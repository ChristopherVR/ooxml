import { savedFillGradient } from './saved-fill-gradient.js';
import { themeLinearGradient } from './theme-gradient.js';
import { themeLineWeight, reportThemeEffects, themeSolidLinePattern } from './theme-line.js';
import { lineCap } from './line-style.js';
import { linePattern } from './line-pattern.js';
import { themeColor, type ThemeResources } from './theme-resolve.js';
import { textBackground } from './text-background.js';
import { textParagraphs, type ParagraphMarker } from './paragraphs.js';
import type { VisioStyle, VisioText, VisioTextRun } from './model.js';
import { transform } from './geometry.js';
import { number, sectionRows, type Cells, type Report, type Sheet } from './sheet.js';

export interface Resources extends ThemeResources {
	colors: Map<string, string>;
	fonts: Map<string, string>;
}
const PALETTE = [
	'#000000',
	'#ffffff',
	'#ff0000',
	'#00ff00',
	'#0000ff',
	'#ffff00',
	'#ff00ff',
	'#00ffff',
	'#800000',
	'#008000',
	'#000080',
	'#808000',
	'#800080',
	'#008080',
	'#c0c0c0',
	'#e6e6e6',
	'#cdcdcd',
	'#b3b3b3',
	'#9a9a9a',
	'#808080',
	'#666666',
	'#4d4d4d',
	'#333333',
	'#1a1a1a',
];
export function color(
	cells: Cells,
	name: string,
	fallback: string,
	resources: Resources,
	report: Report,
	themeCells: Cells = cells,
	gradientRendered = false,
): string {
	const value = cells.get(name)?.value;
	if (value === undefined) return fallback;
	if (/^#[0-9a-f]{6}$/i.test(value)) return value;
	if (/^\d+$/.test(value)) {
		const entry = resources.colors.get(value) ?? PALETTE[Number(value)];
		if (entry && /^#[0-9a-f]{6}$/i.test(entry)) return entry;
	}
	if (value === 'Themed') {
		const resolved = themeColor(themeCells, name, resources, report, gradientRendered);
		if (resolved) return resolved;
	}
	report('unsupported-color', `Cell ${name} uses an unresolved color; a default was used.`);
	return fallback;
}
const opacity = (value: number) => Math.max(0, Math.min(1, 1 - value));
export function shapeStyle(
	sheet: Sheet,
	resources: Resources,
	report: Report,
	width = number(sheet.cells, 'Width', 1),
	height = number(sheet.cells, 'Height', 1),
): VisioStyle {
	const cells = sheet.cells,
		pattern = number(cells, 'FillPattern', 1, report);
	const savedGradient =
		pattern !== 0
			? savedFillGradient(
					sheet,
					width,
					height,
					(stops) => color(stops, 'GradientStopColor', '', resources, report),
					report,
				)
			: undefined;
	const fillGradient =
		savedGradient ??
		(pattern === 1 ? themeLinearGradient(sheet, resources, width, height, report) : undefined);
	if (pattern > 1 && !fillGradient)
		report(
			'unsupported-fill-pattern',
			`Fill pattern ${pattern} is approximated by its foreground color.`,
		);
	if (
		(!fillGradient && number(cells, 'FillGradientEnabled', 0, report)) ||
		number(cells, 'LineGradientEnabled', 0, report)
	)
		report(
			'unsupported-gradient',
			'Gradient fills and strokes are approximated with their foreground colors.',
		);
	reportThemeEffects(cells, resources, report);
	if (number(cells, 'ShdwPattern', 0, report))
		report('unsupported-shadow', 'Shape shadows are not rendered.');
	const cap = lineCap(cells, resources, report);
	return {
		fill:
			pattern === 0
				? 'none'
				: color(cells, 'FillForegnd', '#ffffff', resources, report, cells, !!fillGradient),
		...(fillGradient ? { fillGradient } : {}),
		lineColor: color(cells, 'LineColor', '#000000', resources, report),
		lineWidth:
			themeLineWeight(cells, resources, report) ??
			Math.max(0, number(cells, 'LineWeight', 0.01, report)),
		...(cap === undefined ? {} : { lineCap: cap }),
		...(themeSolidLinePattern(cells, resources) ?? linePattern(cells, report)),
		// Saved stop transparencies are the fill opacity, not a second alpha layer.
		fillOpacity: savedGradient ? 1 : opacity(number(cells, 'FillForegndTrans', 0, report)),
		lineOpacity: opacity(number(cells, 'LineColorTrans', 0, report)),
		startArrow: number(cells, 'BeginArrow', 0, report),
		endArrow: number(cells, 'EndArrow', 0, report),
		startArrowSize: number(cells, 'BeginArrowSize', 2, report),
		endArrowSize: number(cells, 'EndArrowSize', 2, report),
	};
}
function font(cells: Cells, resources: Resources): string {
	const value = cells.get('Font')?.value;
	return (
		(value
			? (resources.fonts.get(value) ??
				(!/^\d+$/.test(value) && value !== 'Themed' ? value.replace(/^"|"$/g, '') : undefined))
			: undefined) || 'Arial'
	);
}
function runStyle(
	cells: Cells,
	resources: Resources,
	report: Report,
	themeCells: Cells,
): Omit<VisioTextRun, 'text'> {
	const bits = number(cells, 'Style', 0, report);
	return {
		fontFamily: font(cells, resources),
		fontSize: Math.max(0.001, number(cells, 'Size', 10 / 72, report)),
		color: color(cells, 'Color', '#000000', resources, report, themeCells),
		bold: !!(bits & 1),
		italic: !!(bits & 2),
		underline: !!(bits & 4),
	};
}
export function shapeText(
	sheet: Sheet,
	node: Element | undefined,
	width: number,
	height: number,
	resources: Resources,
	report: Report,
	consume: (length: number) => void,
	consumeParagraph: () => void,
): VisioText {
	const cells = sheet.cells;
	const characters = new Map(sectionRows(sheet, 'Character').map((row) => [row.index, row.cells]));
	const defaultCells =
		characters.get('0') ?? new Map<string, { value?: string; formula?: string }>();
	let currentCells = defaultCells;
	const runs: VisioTextRun[] = [];
	const paragraphMarkers: ParagraphMarker[] = [];
	let textOffset = 0;
	const append = (text: string) => {
		if (text) {
			consume(text.length);
			runs.push({ text, ...runStyle(currentCells, resources, report, cells) });
			textOffset += text.length;
		}
	};
	if (node && !number(cells, 'HideText', 0, report)) {
		for (const part of Array.from(node.childNodes)) {
			if (part.nodeType === 3 || part.nodeType === 4) append(part.nodeValue ?? '');
			else if (part.nodeType === 1) {
				const element = part as Element;
				if (element.localName === 'cp')
					currentCells = characters.get(element.getAttribute('IX') ?? '0') ?? defaultCells;
				else if (element.localName === 'pp')
					paragraphMarkers.push({ offset: textOffset, index: element.getAttribute('IX') ?? '0' });
				else if (element.localName === 'fld') append(element.textContent ?? '');
				else if (element.localName !== 'pp' && element.localName !== 'tp') {
					report(
						'unsupported-text-element',
						`Text element ${element.localName} is represented as plain text.`,
					);
					append(element.textContent ?? '');
				}
			}
		}
	}
	const paragraph = sectionRows(sheet, 'Paragraph')[0]?.cells ?? new Map();
	const alignment = number(paragraph, 'HorzAlign', 1, report);
	const textWidth = Math.max(0, number(cells, 'TxtWidth', width, report)),
		textHeight = Math.max(0, number(cells, 'TxtHeight', height, report));
	const vertical = number(cells, 'VerticalAlign', 1, report);
	const plainText = runs.map((r) => r.text).join('');
	const defaultRun = runStyle(defaultCells, resources, report, cells);
	const paragraphs = textParagraphs(
		plainText,
		paragraphMarkers,
		sheet,
		defaultRun.fontFamily,
		defaultRun.fontSize,
		resources,
		report,
		runs,
		consumeParagraph,
	);
	return {
		plainText,
		...textBackground(cells, resources, report),
		paragraphs,
		runs,
		...runStyle(defaultCells, resources, report, cells),
		horizontalAlign: alignment === 1 ? 'center' : alignment === 2 ? 'right' : 'left',
		verticalAlign: vertical === 0 ? 'top' : vertical === 2 ? 'bottom' : 'middle',
		width: textWidth,
		height: textHeight,
		transform: transform(
			number(cells, 'TxtPinX', width / 2, report),
			number(cells, 'TxtPinY', height / 2, report),
			number(cells, 'TxtLocPinX', textWidth / 2, report),
			number(cells, 'TxtLocPinY', textHeight / 2, report),
			number(cells, 'TxtAngle', 0, report),
		),
		margins: {
			left: number(cells, 'LeftMargin', 0.04, report),
			right: number(cells, 'RightMargin', 0.04, report),
			top: number(cells, 'TopMargin', 0.04, report),
			bottom: number(cells, 'BottomMargin', 0.04, report),
		},
	};
}
