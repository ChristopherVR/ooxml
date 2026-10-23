// SmartArt node text on the neutral DrawingML text model: reads a node's `dgm:t` (or a cached
// `dsp:txBody`) from the shared DOM into the typed node paragraphs, and the styler that lets
// `projectDiagramNodeText` project them without any product's text types. Word and Excel read
// SmartArt through this; pptx keeps its own object-tree reader and styler for now.
import type { XmlElement } from '../../xml/index';
import { parseDrawingRichTextBody } from '../../drawingml/read-text-body';
import type {
	DrawingBullet,
	DrawingParagraphItem,
	DrawingParagraphProperties,
	DrawingRunProperties,
} from '../../drawingml/text-model';
import { keptChildName } from '../../drawingml/write-text-xml';
import type { DiagramNodeTextParagraph, DiagramNodeTextParagraphItem } from '../model/node';
import {
	projectDiagramNodeText,
	type DiagramTextSegment,
	type DiagramTextStyler,
	type ProjectDiagramNodeTextOptions,
} from './node-text-projection';

/** A projected segment's style on the neutral model: the paragraph's and the run's properties. */
export interface DiagramNeutralTextStyle {
	paragraph: DrawingParagraphProperties;
	run: DrawingRunProperties;
}

export type DiagramNeutralTextParagraph = DiagramNodeTextParagraph<DiagramNeutralTextStyle>;
export type DiagramNeutralTextSegment = DiagramTextSegment<DiagramNeutralTextStyle, DrawingBullet>;

function item(value: DrawingParagraphItem): DiagramNodeTextParagraphItem<DiagramNeutralTextStyle> {
	switch (value.kind) {
		case 'run':
			return {
				kind: 'run',
				run: { text: value.text, ...(value.properties ? { properties: value.properties } : {}) },
			};
		case 'break':
			return { kind: 'break', ...(value.properties ? { properties: value.properties } : {}) };
		case 'field':
			return {
				kind: 'field',
				id: value.id,
				...(value.type ? { fieldType: value.type } : {}),
				text: value.text,
				...(value.properties ? { properties: value.properties } : {}),
				...(value.paragraphProperties ? { paragraphProperties: value.paragraphProperties } : {}),
			};
		default:
			return { kind: 'raw', name: keptChildName(value.xml), value: value.xml };
	}
}

/**
 * Reads the paragraphs of a node's text body (`dgm:t`, `dsp:txBody`) into the typed node model,
 * neutral properties in the `properties` slots; `undefined` without a body or paragraphs.
 */
export function readDiagramNodeParagraphs(
	body: XmlElement | undefined,
): DiagramNeutralTextParagraph[] | undefined {
	const parsed = parseDrawingRichTextBody(body);
	if (!parsed?.paragraphs.length) return undefined;
	return parsed.paragraphs.map((paragraph) => ({
		...(paragraph.properties ? { properties: paragraph.properties } : {}),
		items: paragraph.items.map(item),
		...(paragraph.endProperties ? { endProperties: paragraph.endProperties } : {}),
	}));
}

/**
 * The neutral styler. A paragraph's properties overlay the fallback's (its `a:defRPr` is not
 * applied to runs: PowerPoint ignores a paragraph-level `defRPr` when drawing), a run's
 * properties overlay its paragraph's run style, and the empty segment of a paragraph without
 * runs takes `a:endParaRPr`.
 */
export const neutralDiagramTextStyler: DiagramTextStyler<DiagramNeutralTextStyle, DrawingBullet> = {
	paragraphStyle: (paragraph, fallback) => ({
		paragraph: { ...fallback.paragraph, ...paragraph.properties },
		run: fallback.run,
	}),
	runStyle: (value, paragraphStyle) => ({
		paragraph: paragraphStyle.paragraph,
		run: { ...paragraphStyle.run, ...value.properties },
	}),
	endParagraphStyle: (paragraph, paragraphStyle) => ({
		paragraph: paragraphStyle.paragraph,
		run: { ...paragraphStyle.run, ...paragraph.endProperties },
	}),
	paragraphLevel: (paragraph) => paragraph.properties?.level ?? 0,
	explicitBullet: (paragraph) => paragraph.properties?.bullet,
	autoBullet: () => ({ kind: 'char', char: '•' }),
};

/** Projects a node read through {@link readDiagramNodeParagraphs} into neutral text segments. */
export function projectDiagramNodeTextNeutral(
	node: { text: string; paragraphs?: DiagramNeutralTextParagraph[] },
	fallbackStyle: DiagramNeutralTextStyle = { paragraph: {}, run: {} },
	options: ProjectDiagramNodeTextOptions = {},
): DiagramNeutralTextSegment[] {
	return projectDiagramNodeText(node, neutralDiagramTextStyler, fallbackStyle, options);
}
