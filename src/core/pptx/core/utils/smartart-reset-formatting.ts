import type { PptxSmartArtData, PptxSmartArtNode } from '../types';
import type { PptxSmartArtTextParagraph, PptxSmartArtTextRun } from '../types/smart-art';
import { relayoutSmartArt } from './smartart-relayout';

const RUN_METADATA = new Set([
	'@_lang',
	'@_altLang',
	'@_dirty',
	'@_err',
	'@_noProof',
	'extLst',
	'hlinkClick',
	'hlinkMouseOver',
]);

/** Keep linguistic metadata and extension data when dropping direct run formatting. */
export function isSmartArtRunMetadata(name: string): boolean {
	return RUN_METADATA.has(name.split(':').at(-1) ?? name);
}

function resetSmartArtRunProperties(
	properties: Record<string, unknown> | undefined,
): Record<string, unknown> | undefined {
	if (!properties) return undefined;
	return Object.fromEntries(
		Object.entries(properties).filter(([name]) => isSmartArtRunMetadata(name)),
	);
}

function resetParagraphProperties(
	properties: Record<string, unknown> | undefined,
): Record<string, unknown> | undefined {
	if (!properties) return undefined;
	const keep = new Set(['@_lvl', 'buNone', 'buChar', 'buAutoNum', 'buBlip', 'extLst']);
	return Object.fromEntries(
		Object.entries(properties).filter(([name]) => keep.has(name.split(':').at(-1) ?? name)),
	);
}

function resetRun(run: PptxSmartArtTextRun): PptxSmartArtTextRun {
	const { rPr, style: _style, properties: _properties, ...content } = run;
	const reset = resetSmartArtRunProperties(rPr);
	return { ...content, ...(reset && { rPr: reset }) };
}

function resetParagraph(paragraph: PptxSmartArtTextParagraph): PptxSmartArtTextParagraph {
	const {
		pPr,
		endParaRPr,
		endParaStyle: _endParaStyle,
		properties: _properties,
		endProperties: _endProperties,
		...content
	} = paragraph;
	const reset = resetParagraphProperties(pPr);
	const end = resetSmartArtRunProperties(endParaRPr);
	return {
		...content,
		...(reset && { pPr: reset }),
		...(end && { endParaRPr: end }),
		items: paragraph.items.map((item) => {
			if (item.kind === 'run') return { ...item, run: resetRun(item.run) };
			if (item.kind !== 'break' && item.kind !== 'field') return item;
			const { rPr, style: _style, properties: _properties, ...content } = item;
			const runProperties = resetSmartArtRunProperties(rPr);
			if (content.kind === 'break')
				return { ...content, ...(runProperties && { rPr: runProperties }) };
			const { pPr, paragraphProperties: _paragraphProperties, ...field } = content;
			const fieldParagraph = resetParagraphProperties(pPr);
			return {
				...field,
				...(runProperties && { rPr: runProperties }),
				...(fieldParagraph && { pPr: fieldParagraph }),
			};
		}),
	};
}

function resetNode(node: PptxSmartArtNode): PptxSmartArtNode {
	const {
		style: _style,
		customLayout: _customLayout,
		coherent3DOff: _coherent3DOff,
		...content
	} = node;
	return {
		...content,
		runs: node.runs?.map(resetRun),
		paragraphs: node.paragraphs?.map(resetParagraph),
		children: node.children?.map(resetNode),
	};
}

/** Clear node overrides and regenerate the drawing from the current diagram choices. */
export function resetSmartArtFormatting(
	data: PptxSmartArtData,
	box: { width: number; height: number },
	themeColorMap?: Readonly<Record<string, string>>,
): PptxSmartArtData {
	const next: PptxSmartArtData = {
		...data,
		nodes: data.nodes.map(resetNode),
		drawingShapes: [],
		drawingDirty: true,
		clearDirectNodeFormatting: true,
	};
	return {
		...next,
		drawingShapes: relayoutSmartArt(
			next,
			box.width,
			box.height,
			themeColorMap && { ...themeColorMap },
		),
	};
}
