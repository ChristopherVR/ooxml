import type { DiagramDrawingShape } from 'ooxml-core/diagram';
import type { DrawingTextRun } from 'ooxml-core/drawingml';
import { designTextWidthEm, wrapStyledRuns } from 'ooxml-core/text';
import {
	colorToCss,
	type SchemeColors,
	type SchemeFonts,
	type SmartArtRenderReport,
} from './smartart-svg';

const NS = 'http://www.w3.org/2000/svg';

function runStyle(
	element: SVGElement,
	run: DrawingTextRun,
	fonts: SchemeFonts,
	scheme: SchemeColors,
	report: SmartArtRenderReport,
): void {
	if (run.sizePt !== undefined) element.setAttribute('font-size', String((run.sizePt * 96) / 72));
	if (run.bold !== undefined) element.setAttribute('font-weight', run.bold ? '700' : '400');
	if (run.italic !== undefined)
		element.setAttribute('font-style', run.italic ? 'italic' : 'normal');
	if (run.underline !== undefined)
		element.setAttribute('text-decoration', run.underline ? 'underline' : 'none');
	if (run.color)
		element.setAttribute('fill', colorToCss(run.color, scheme, report) ?? 'currentColor');
	const typeface = run.typeface;
	const font = typeface === '+mj-lt' ? fonts.major : typeface === '+mn-lt' ? fonts.minor : typeface;
	if (font) element.setAttribute('font-family', font);
}

/** Paint every paragraph and run, retaining authored styles instead of flattening to one label. */
export function diagramTextElement(
	doc: Document,
	shape: DiagramDrawingShape,
	scheme: SchemeColors,
	fonts: SchemeFonts,
	report: SmartArtRenderReport,
): SVGElement {
	const frame = shape.textFrame ?? shape.frame;
	const ox = (frame.x - shape.frame.x) / 9525;
	const oy = (frame.y - shape.frame.y) / 9525;
	const w = frame.width / 9525;
	const h = frame.height / 9525;
	const sourceParagraphs = shape.text?.paragraphs ?? [];
	const insets = shape.text?.insetsEmu;
	const contentX = ox + (insets?.left ?? 0) / 9525;
	const contentY = oy + (insets?.top ?? 0) / 9525;
	const contentWidth = Math.max(1, w - ((insets?.left ?? 0) + (insets?.right ?? 0)) / 9525);
	const contentHeight = Math.max(1, h - ((insets?.top ?? 0) + (insets?.bottom ?? 0)) / 9525);
	const paragraphs = sourceParagraphs.flatMap((paragraph, paragraphIndex) => {
		const runs = paragraph.runs.map((run) => ({ ...paragraph.defaultProperties, ...run }));
		const measure = (content: string, run: DrawingTextRun): number => {
			const family =
				run.typeface === '+mj-lt'
					? fonts.major
					: run.typeface === '+mn-lt'
						? fonts.minor
						: run.typeface;
			const em = designTextWidthEm(content, family ?? fonts.minor ?? 'Arial', run.bold);
			if (em === undefined) {
				const issue = `font-metrics:${family ?? fonts.minor ?? 'Arial'}`;
				if (!report.approximatedText?.includes(issue))
					report.approximatedText = [...(report.approximatedText ?? []), issue];
			}
			return ((em ?? [...content].length * 0.5) * (run.sizePt ?? 12) * 96) / 72;
		};
		const maxWidth = shape.body?.bodyProperties?.wrap === 'none' ? Infinity : contentWidth;
		const plain = runs.map((run) => run.text).join('');
		let offset = 0;
		const lines = wrapStyledRuns(runs, maxWidth, measure, { breakWords: true });
		return lines.map((line, index) => {
			const content = line.map((run) => run.text).join('');
			const start = plain.indexOf(content, offset);
			const separator = start >= offset ? plain.slice(offset, start) : '';
			offset = Math.max(offset, start + content.length);
			return {
				...paragraph,
				runs: line,
				separatorBefore: index === 0 && paragraphIndex > 0 ? '\n' : separator,
				first: index === 0,
				last: index === lines.length - 1,
			};
		});
	});
	const first = paragraphs[0]?.runs[0];
	const size = ((first?.sizePt ?? 12) * 96) / 72;
	const text = doc.createElementNS(NS, 'text');
	text.setAttribute('x', String(ox + w / 2));
	text.setAttribute('y', String(oy + h / 2));
	text.setAttribute('text-anchor', 'middle');
	text.setAttribute('dominant-baseline', 'central');
	text.setAttribute('font-size', String(size));
	text.setAttribute(
		'fill',
		colorToCss(first?.color ?? shape.style?.font?.color, scheme, report) ?? 'currentColor',
	);
	const ref = shape.style?.font?.fontIndex;
	const inherited = ref === 'major' ? fonts.major : ref === 'minor' ? fonts.minor : undefined;
	if (inherited) text.setAttribute('font-family', inherited);
	if (first) runStyle(text, first, fonts, scheme, report);
	const body = shape.body?.bodyProperties;
	if (
		body &&
		((body.vertical && body.vertical !== 'horz') || body.rotation || (body.columnCount ?? 1) > 1)
	)
		report.approximatedText = [
			...new Set([...(report.approximatedText ?? []), 'vertical-or-multicolumn']),
		];
	if (
		shape.body?.paragraphs.some(
			(paragraph) =>
				(paragraph.properties?.bullet && paragraph.properties.bullet.kind !== 'none') ||
				paragraph.properties?.tabStops?.length,
		)
	)
		report.approximatedText = [...new Set([...(report.approximatedText ?? []), 'bullets-or-tabs'])];
	const spacing = (
		value: { unit: 'points' | 'percent'; value: number } | undefined,
		nominal: number,
	) => (value ? (value.unit === 'points' ? (value.value * 96) / 72 : value.value * nominal) : 0);
	const heights = paragraphs.map((paragraph) => {
		const nominal = Math.max(size, ...paragraph.runs.map((run) => ((run.sizePt ?? 12) * 96) / 72));
		return {
			line: paragraph.lineSpacing ? spacing(paragraph.lineSpacing, nominal) : nominal * 1.2,
			before: paragraph.first ? spacing(paragraph.spaceBefore, nominal) : 0,
			after: paragraph.last ? spacing(paragraph.spaceAfter, nominal) : 0,
		};
	});
	const totalHeight = heights.reduce(
		(sum, value) => sum + value.line + value.before + value.after,
		0,
	);
	const anchor = shape.text?.anchor;
	let y =
		anchor === 't'
			? contentY
			: anchor === 'b'
				? contentY + contentHeight - totalHeight
				: contentY + (contentHeight - totalHeight) / 2;
	paragraphs.forEach((paragraph, index) => {
		if (paragraph.separatorBefore) text.append(doc.createTextNode(paragraph.separatorBefore));
		const line = doc.createElementNS(NS, 'tspan');
		const align = paragraph.align;
		const x =
			align === 'l'
				? contentX
				: align === 'r'
					? contentX + contentWidth
					: contentX + contentWidth / 2;
		const height = heights[index] ?? { line: size, before: 0, after: 0 };
		y += height.before;
		line.setAttribute('x', String(x));
		line.setAttribute('y', String(y + height.line / 2));
		line.setAttribute('text-anchor', align === 'l' ? 'start' : align === 'r' ? 'end' : 'middle');
		for (const run of paragraph.runs) {
			const span = doc.createElementNS(NS, 'tspan');
			span.textContent = run.text;
			runStyle(
				span,
				{ bold: false, italic: false, underline: false, ...paragraph.defaultProperties, ...run },
				fonts,
				scheme,
				report,
			);
			line.append(span);
		}
		text.append(line);
		y += height.line + height.after;
	});
	return text;
}
