import type { DrawingTextBody, DrawingTextParagraph, DrawingTextRun } from '../drawingml/types';
import type { DiagramHostNode } from './edit/node-edit';
import { projectDiagramNodeTextNeutral } from './text/node-text-neutral';

/** Project edited and folded node text through the same neutral styler as the text pane. */
export function diagramPreviewText(
	nodes: readonly DiagramHostNode[],
	fontSize: number,
): DrawingTextBody {
	const paragraphs: DrawingTextParagraph[] = [];
	for (const node of nodes) {
		let paragraph: DrawingTextParagraph = { runs: [] };
		paragraphs.push(paragraph);
		for (const segment of projectDiagramNodeTextNeutral(node)) {
			if (segment.isParagraphBreak) {
				paragraph = { runs: [] };
				paragraphs.push(paragraph);
				continue;
			}
			const properties = segment.style.run;
			const run: DrawingTextRun = {
				text: segment.text,
				sizePt: properties.sizePt ?? fontSize * 0.75,
				...(properties.bold !== undefined ? { bold: properties.bold } : {}),
				...(properties.italic !== undefined ? { italic: properties.italic } : {}),
				...(properties.underline ? { underline: properties.underline !== 'none' } : {}),
				...(properties.fill?.kind === 'solid' ? { color: properties.fill.color } : {}),
				...(properties.latin?.typeface ? { typeface: properties.latin.typeface } : {}),
			};
			const align = segment.style.paragraph.align;
			if (align) paragraph.align = align;
			paragraph.runs.push(run);
		}
	}
	return {
		anchor: 'ctr',
		paragraphs,
		text: paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join('\n'),
	};
}
