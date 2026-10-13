import type {
	PptxElement,
	PptxSmartArtNode,
	PptxTableCellTextRun,
	TextSegment,
	XmlObject,
} from '../../index';
import { hasTextProperties, updateSmartArtNodeText } from '../../index';
import { ensureArray, getTblFromRawXml } from '../../core/core/runtime/table-structural-helpers';
import { updateCellTextInRawXml } from '../../core/core/runtime/table-cell-rawxml-ops';
import { flattenCellTxBodyText } from '../../core/core/runtime/table-cell-text-xml';
import { reconcileSmartArtTextParagraphs } from '../../core/core/runtime/smartart-text-reconciliation';
import type { FindResult } from './find-replace';
import { isFindBreak, replaceFindRuns, smartArtNodes } from './find-text-targets';
import { remapTextToSegments } from './remap-text';
import { chartLabelsAreEditable } from './chart-label-editability';

function patchTableXml(
	element: PptxElement,
	rowIndex: number,
	colIndex: number,
	original: readonly PptxTableCellTextRun[],
	runs: readonly PptxTableCellTextRun[],
): XmlObject | undefined {
	if (!element.rawXml) return undefined;
	const rawXml = structuredClone(element.rawXml);
	const table = getTblFromRawXml(rawXml);
	const row = ensureArray(table?.['a:tr'] as XmlObject | XmlObject[] | undefined)[rowIndex];
	const cell = ensureArray(row?.['a:tc'] as XmlObject | XmlObject[] | undefined)[colIndex];
	const paragraphs = ensureArray(
		(cell?.['a:txBody'] as XmlObject | undefined)?.['a:p'] as XmlObject | XmlObject[] | undefined,
	);
	let paragraphIndex = 0,
		runIndex = 0,
		fieldIndex = 0;
	for (let index = 0; index < original.length; index++) {
		const previous = original[index]!;
		if (previous.isParagraphBreak) {
			paragraphIndex++;
			runIndex = 0;
			fieldIndex = 0;
			continue;
		}
		if (previous.isLineBreak) continue;
		const paragraph = paragraphs[paragraphIndex];
		const key = previous.isField ? 'a:fld' : 'a:r';
		const xmlRun = ensureArray(paragraph?.[key] as XmlObject | XmlObject[] | undefined)[
			previous.isField ? fieldIndex++ : runIndex++
		];
		if (!xmlRun) return undefined;
		if (previous.text !== runs[index]!.text) xmlRun['a:t'] = runs[index]!.text;
	}
	return rawXml;
}

function updateNode(nodes: PptxSmartArtNode[], nodeId: string, text: string): PptxSmartArtNode[] {
	return nodes.map((node) => {
		if (node.id === nodeId) {
			const originalParagraphs =
				node.paragraphs ??
				(node.runs?.length
					? [{ items: node.runs.map((run) => ({ kind: 'run' as const, run })) }]
					: undefined);
			const paragraphs = originalParagraphs
				? reconcileSmartArtTextParagraphs(originalParagraphs, text)
				: undefined;
			return { ...node, text, ...(paragraphs ? { paragraphs } : {}) };
		}
		if (!node.children) return node;
		const children = updateNode(node.children, nodeId, text);
		return children.every((child, index) => child === node.children![index])
			? node
			: { ...node, children };
	});
}

/** Apply one match at its leaf without mutating the surrounding group or sibling objects. */
function replaceLeaf(element: PptxElement, match: FindResult, replacement: string): PptxElement {
	if (
		(match.chartSeriesIndex !== undefined || match.chartCategoryIndex !== undefined) &&
		element.type === 'chart' &&
		element.chartData
	) {
		const data = element.chartData;
		if (!chartLabelsAreEditable(element))
			throw new Error('Chart labels require editing in PowerPoint or their source workbook');
		const text =
			match.chartSeriesIndex !== undefined
				? data.series[match.chartSeriesIndex]?.name
				: data.categories[match.chartCategoryIndex!];
		if (text === undefined) return element;
		const next = replaceFindRuns(
			[{ text }],
			match.segmentIndex,
			match.startOffset,
			match.length,
			replacement,
		);
		if (!next) return element;
		return {
			...element,
			chartData: {
				...data,
				dataLabelTextOnly: true,
				...(match.chartSeriesIndex !== undefined
					? {
							series: data.series.map((series, index) =>
								index === match.chartSeriesIndex ? { ...series, name: next[0]!.text } : series,
							),
						}
					: {
							categories: data.categories.map((category, index) =>
								index === match.chartCategoryIndex ? next[0]!.text : category,
							),
						}),
			},
		};
	}
	if (match.chartTitle && element.type === 'chart' && element.chartData) {
		const data = element.chartData;
		if (data.titleTextEditable === false)
			throw new Error('This chart title requires editing in PowerPoint or its source workbook');
		const runs = data.titleRuns?.length ? data.titleRuns : [{ text: data.title ?? '' }];
		const next = replaceFindRuns(
			runs,
			match.segmentIndex,
			match.startOffset,
			match.length,
			replacement,
		);
		if (!next) return element;
		return {
			...element,
			chartData: {
				...data,
				title: next[0]?.text ?? '',
				titleRuns: next,
				titleTextOnly: true,
				titleTextOnlyStyles:
					data.titleTextOnlyStyles ?? runs.map(({ text: _text, ...style }) => style),
				style: { ...data.style, hasTitle: next.some((run) => run.text.length > 0) },
			},
		};
	}
	if (match.tableCell && element.type === 'table' && element.tableData) {
		const { rowIndex, colIndex } = match.tableCell;
		const row = element.tableData.rows[rowIndex],
			cell = row?.cells[colIndex];
		if (!cell) return element;
		const original = cell.textRuns?.length ? cell.textRuns : [{ text: cell.text }];
		const runs = replaceFindRuns(
			original,
			match.segmentIndex,
			match.startOffset,
			match.length,
			replacement,
		);
		if (!runs) return element;
		let text = runs
			.map((run) => (run.isParagraphBreak ? '\n' : run.isLineBreak ? '' : run.text))
			.join('');
		let rawXml = cell.textRuns?.length
			? patchTableXml(element, rowIndex, colIndex, original, runs)
			: updateCellTextInRawXml(element, rowIndex, colIndex, text);
		if (rawXml) {
			const table = getTblFromRawXml(rawXml);
			const xmlRow = ensureArray(table?.['a:tr'] as XmlObject | XmlObject[] | undefined)[rowIndex];
			const xmlCell = ensureArray(xmlRow?.['a:tc'] as XmlObject | XmlObject[] | undefined)[
				colIndex
			];
			text = flattenCellTxBodyText(xmlCell?.['a:txBody'] as XmlObject | undefined, (value) =>
				ensureArray(value as XmlObject | XmlObject[] | undefined),
			);
		}
		const rows = [...element.tableData.rows],
			cells = [...row!.cells];
		cells[colIndex] = { ...cell, text, ...(cell.textRuns ? { textRuns: runs } : {}) };
		rows[rowIndex] = { ...row!, cells };
		return { ...element, tableData: { ...element.tableData, rows }, ...(rawXml ? { rawXml } : {}) };
	}
	if (match.smartArtNodeId && element.type === 'smartArt' && element.smartArtData) {
		const data = element.smartArtData;
		const nodes = [...smartArtNodes(data.nodes)],
			node = nodes.find((candidate) => candidate.id === match.smartArtNodeId);
		if (!node || match.startOffset < 0 || match.startOffset + match.length > node.text.length)
			return element;
		const text =
			node.text.slice(0, match.startOffset) +
			replacement +
			node.text.slice(match.startOffset + match.length);
		const next = updateSmartArtNodeText({ ...data, nodes }, node.id, text);
		const drawingShapes = next.drawingShapes?.map((shape, index) => {
			if (shape === data.drawingShapes?.[index] || !shape.textSegments) return shape;
			return {
				...shape,
				textSegments: remapTextToSegments(text, shape.textSegments, { fontSize: shape.fontSize }),
			};
		});
		return {
			...element,
			smartArtData: { ...next, drawingShapes, nodes: updateNode(data.nodes, node.id, text) },
		};
	}
	if (!hasTextProperties(element)) return element;
	const original: TextSegment[] = element.textSegments?.length
		? element.textSegments
		: element.text
			? [{ text: element.text, style: element.textStyle ?? {} }]
			: [];
	const segments = replaceFindRuns(
		original,
		match.segmentIndex,
		match.startOffset,
		match.length,
		replacement,
	);
	if (!segments) return element;
	return {
		...element,
		textSegments: segments,
		text: segments.map((run) => (isFindBreak(run) ? '\n' : run.text)).join(''),
	} as PptxElement;
}

export function replaceFindTarget(
	element: PptxElement,
	match: FindResult,
	replacement: string,
	path = match.elementPath ?? [],
): PptxElement {
	if (!path.length) return replaceLeaf(element, match, replacement);
	if (element.type !== 'group') return element;
	const index = path[0]!,
		child = element.children[index];
	if (!child) return element;
	const next = replaceFindTarget(child, match, replacement, path.slice(1));
	if (next === child) return element;
	const children = [...element.children];
	children[index] = next;
	return { ...element, children };
}
