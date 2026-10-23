// Projects a SmartArt node's typed paragraphs into the flat text-segment list renderers lay out:
// one segment per run, field, line break and tab, paragraph metadata (bullet, level, end run
// properties) on the first segment of each paragraph, and a paragraph-break segment between
// paragraphs. Moved from `pptx/core/utils/smartart-node-text-projection.ts`.
//
// The structure is shared; what a "style" is belongs to the reader. pptx passes a styler built on
// its object-tree text parsers (its `TextStyle`); the DOM readers use the neutral styler in
// `node-text-neutral.ts`, over the `drawingml` text model.
import type { DrawingRunProperties } from '../../drawingml/text-model';
import type { DiagramNode, DiagramNodeTextParagraph } from '../model/node';
import { reconcileDiagramTextParagraphs } from './node-text-reconciliation';

/** One projected segment. `S` is the reader's style, `B` its bullet description. */
export interface DiagramTextSegment<S, B> {
	text: string;
	style: S;
	/** On the first segment of a paragraph: its bullet, explicit or synthesized. */
	bulletInfo?: B;
	/** On the first segment of a paragraph: its outline level (0-8). */
	paragraphLevel?: number;
	/** On the first segment of a paragraph: the raw `a:endParaRPr` (object-tree readers). */
	endParaRunProperties?: Record<string, unknown>;
	fieldType?: string | undefined;
	fieldGuid?: string | undefined;
	fieldParagraphPropertiesXml?: Record<string, unknown> | undefined;
	isLineBreak?: true;
	breakRunProperties?: Record<string, unknown> | undefined;
	isParagraphBreak?: boolean;
}

/** How a reader turns paragraph and run properties into its style and bullet types. */
export interface DiagramTextStyler<S, B> {
	/** The style a paragraph's runs inherit: `fallback` overlaid with the paragraph properties. */
	paragraphStyle(paragraph: DiagramNodeTextParagraph<S>, fallback: S): S;
	/** A run's style (`item` is the run, field or break carrying `rPr`/`properties`/`style`). */
	runStyle(item: DiagramTextStyledItem<S>, paragraphStyle: S): S;
	/** The style of the empty segment standing for a paragraph with no text children. */
	endParagraphStyle(paragraph: DiagramNodeTextParagraph<S>, paragraphStyle: S): S;
	/** The paragraph's outline level, 0 when absent. */
	paragraphLevel(paragraph: DiagramNodeTextParagraph<S>): number;
	/** The paragraph's own bullet opinion (`a:buNone` included); `undefined` for none. */
	explicitBullet(paragraph: DiagramNodeTextParagraph<S>, paragraphIndex: number): B | undefined;
	/** The bullet synthesized for `dgm:bulletEnabled` levels without an opinion of their own. */
	autoBullet(): B;
}

/** The style-carrying part of a run, field or break item. */
export interface DiagramTextStyledItem<S> {
	rPr?: Record<string, unknown>;
	style?: S;
	properties?: DrawingRunProperties;
}

/** Options refining how {@link projectDiagramNodeText} builds paragraph metadata. */
export interface ProjectDiagramNodeTextOptions {
	/** `presLayoutVars.bulletEnabled` for the diagram this node belongs to. */
	bulletEnabled?: boolean;
}

/**
 * `dgm:presLayoutVars/dgm:bulletEnabled` (CT_BulletEnabled) auto-bullets a node's subordinate
 * outline levels even when the paragraph declares no bullet markup (the "Vertical Bullet
 * List"-family layout definitions never write `a:buChar` on their item template). An explicit
 * opinion on the paragraph, `a:buNone` included, always wins; only a level >= 1 (the node's own
 * title line never bullets) with no opinion gets the synthesized marker.
 */
function paragraphMetadata<S, B>(
	styler: DiagramTextStyler<S, B>,
	paragraph: DiagramNodeTextParagraph<S>,
	paragraphIndex: number,
	bulletEnabled: boolean,
): Partial<DiagramTextSegment<S, B>> {
	const level = styler.paragraphLevel(paragraph);
	const explicitBullet = styler.explicitBullet(paragraph, paragraphIndex);
	const bulletInfo =
		explicitBullet ?? (bulletEnabled && level >= 1 ? styler.autoBullet() : undefined);
	return {
		...(bulletInfo ? { bulletInfo } : {}),
		paragraphLevel: level,
		...(paragraph.endParaRPr ? { endParaRunProperties: paragraph.endParaRPr } : {}),
	};
}

/** Projects a node's paragraphs (reconciled with its flat text) into text segments. */
export function projectDiagramNodeText<S, B>(
	node: Pick<DiagramNode<S>, 'text' | 'paragraphs'>,
	styler: DiagramTextStyler<S, B>,
	fallbackStyle: S,
	options: ProjectDiagramNodeTextOptions = {},
): DiagramTextSegment<S, B>[] {
	if (!node.paragraphs?.length) {
		return [{ text: node.text, style: fallbackStyle }];
	}
	const paragraphs = reconcileDiagramTextParagraphs(node.paragraphs, node.text);
	const segments: DiagramTextSegment<S, B>[] = [];
	for (let paragraphIndex = 0; paragraphIndex < paragraphs.length; paragraphIndex++) {
		const paragraph = paragraphs[paragraphIndex]!;
		const pStyle = styler.paragraphStyle(paragraph, fallbackStyle);
		const metadata = paragraphMetadata(
			styler,
			paragraph,
			paragraphIndex,
			options.bulletEnabled ?? false,
		);
		let first = true;
		const push = (segment: DiagramTextSegment<S, B>): void => {
			segments.push(first ? { ...segment, ...metadata } : segment);
			first = false;
		};
		for (const item of paragraph.items) {
			if (item.kind === 'run') {
				push({ text: item.run.text, style: styler.runStyle(item.run, pStyle) });
			} else if (item.kind === 'field') {
				push({
					text: item.text,
					style: styler.runStyle(item, pStyle),
					fieldType: item.fieldType,
					fieldGuid: item.id,
					fieldParagraphPropertiesXml: item.pPr,
				});
			} else if (item.kind === 'break') {
				push({
					text: '\n',
					style: styler.runStyle(item, pStyle),
					isLineBreak: true,
					breakRunProperties: item.rPr,
				});
			} else if (item.kind === 'tab') {
				push({ text: '\t', style: pStyle });
			}
		}
		if (first) {
			push({ text: '', style: styler.endParagraphStyle(paragraph, pStyle) });
		}
		if (paragraphIndex < paragraphs.length - 1) {
			segments.push({
				text: '',
				style: styler.endParagraphStyle(paragraph, pStyle),
				isParagraphBreak: true,
			});
		}
	}
	return segments;
}
