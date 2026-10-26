// Explicit rich title edits preserve paragraph properties and unchanged run XML.
import { NS, first, elements } from '../../xml';
import type { ChartSpace } from '../model';
import { chartTextBodyFromXml } from '../write-shape';
import { innerXml } from '../xml-fragment';
import { editChartSpace } from './draft';
import { newChartTitle } from './title';
import { chartTextBodyElement, editChartTextRunProperties, type ChartTextEdit } from './text';

export interface ChartTitleRunEdit {
	text: string;
	/** Only supplied formatting fields change; absent properties keep their authored markup. */
	format?: ChartTextEdit;
	/** Remove explicit boolean run attributes to restore paragraph inheritance. */
	inherit?: readonly ('bold' | 'italic')[];
}
export function setChartTitleRuns(
	space: ChartSpace,
	runs: readonly ChartTitleRunEdit[],
): ChartSpace {
	return editChartSpace(space, (draft) => {
		const title = draft.title ?? newChartTitle('');
		const root = chartTextBodyElement(title.tx?.rich);
		const paragraphs = elements(root).filter(
			(node) => node.namespaceURI === NS.a && node.localName === 'p',
		);
		let last = paragraphs.at(-1);
		if (!last) {
			last = root.ownerDocument.createElementNS(NS.a, 'a:p');
			root.appendChild(last);
		}
		const authored = paragraphs.flatMap((paragraph) =>
			elements(paragraph).filter((node) => node.namespaceURI === NS.a && node.localName === 'r'),
		);
		for (const [index, input] of runs.entries()) {
			let run = authored[index];
			if (!run) {
				run = root.ownerDocument.createElementNS(NS.a, 'a:r');
				const end = first(last, 'endParaRPr', NS.a);
				last.insertBefore(run, end ?? null);
			}
			let text = first(run, 't', NS.a);
			if (!text) {
				text = root.ownerDocument.createElementNS(NS.a, 'a:t');
				run.appendChild(text);
			}
			text.textContent = input.text;
			if (input.format || input.inherit?.length) {
				let props = first(run, 'rPr', NS.a);
				if (!props) {
					props = root.ownerDocument.createElementNS(NS.a, 'a:rPr');
					run.insertBefore(props, run.firstChild);
				}
				if (input.format) editChartTextRunProperties(props, input.format);
				for (const key of input.inherit ?? []) props.removeAttribute(key === 'bold' ? 'b' : 'i');
			}
		}
		for (const run of authored.slice(runs.length)) run.parentNode?.removeChild(run);
		const body = chartTextBodyFromXml(innerXml(root));
		title.tx = { rich: body, text: body.text };
		title.text = body.text;
		draft.title = title;
		draft.autoTitleDeleted = false;
	});
}
