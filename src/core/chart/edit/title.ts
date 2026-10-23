// Chart and axis title edits. Text edits follow Excel's semantics (xlsx `write/chart-patch.ts`):
// an existing rich title keeps its formatting, its first run takes the new text and the other runs
// go; a new title is one run under the paragraph properties Excel writes, so Office still resolves
// the inherited title font. Removing the chart title also sets `c:autoTitleDeleted`, or Office would
// draw an automatic title in its place.
import { NS, elements, parseXml } from '../../xml/index';
import type { ChartSpace } from '../model';
import type { ChartTextBody, ChartTitle } from '../model-series';
import { chartTextBodyFromXml } from '../write-shape';
import { escapeText } from '../write-util';
import { innerXml } from '../xml-fragment';
import { editChartSpace } from './draft';

/** The children of `c:rich` for a new one-run title (`<a:bodyPr/>`, `<a:lstStyle/>`, one `a:p`). */
export const chartTitleRichXml = (text: string): string =>
	`<a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>${escapeText(text)}</a:t></a:r></a:p>`;

/** A new title in the neutral model, written as Excel writes a new chart title. */
export function newChartTitle(text: string): ChartTitle {
	const rich = chartTextBodyFromXml(chartTitleRichXml(text));
	return { tx: { rich, text: rich.text }, text: rich.text, overlay: false };
}

/** A rich body whose first run reads `text`, other runs removed; `undefined` if it has no run. */
function retextBody(body: ChartTextBody, text: string): ChartTextBody | undefined {
	if (body.sourceXml !== undefined) {
		const root = parseXml(
			`<c:rich xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}">${body.sourceXml}</c:rich>`,
			{ label: 'Chart title' },
		).documentElement;
		const runs = [...root.getElementsByTagNameNS(NS.a, 'r')];
		const first = runs[0]?.getElementsByTagNameNS(NS.a, 't')[0];
		if (!first) return undefined;
		first.textContent = text;
		for (const run of runs.slice(1)) run.parentNode?.removeChild(run);
		return chartTextBodyFromXml(innerXml(root));
	}
	const run = body.paragraphs.flatMap((paragraph) => paragraph.runs)[0];
	if (!run) return undefined;
	let placed = false;
	const paragraphs = body.paragraphs.map((paragraph) => {
		const runs = !placed && paragraph.runs.includes(run) ? [{ ...run, text }] : [];
		placed ||= runs.length > 0;
		return { ...paragraph, runs };
	});
	return {
		...body,
		paragraphs,
		text: paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join('\n'),
	};
}

/** `title` (or a new title) reading `text`, keeping layout, overlay, shape and text properties. */
export function chartTitleWithText(title: ChartTitle | undefined, text: string): ChartTitle {
	if (!title) return newChartTitle(text);
	const rich = title.tx?.reference ? undefined : title.tx?.rich;
	const body = rich ? retextBody(rich, text) : undefined;
	const fresh = newChartTitle(text);
	const next: ChartTitle = {
		...structuredClone(title),
		tx: body ? { rich: body, text: body.text } : fresh.tx!,
	};
	next.text = next.tx!.text!;
	return next;
}

/**
 * Sets the chart title text, or removes the title (`undefined`), as Excel's chart title box does.
 * Semantics: Excel (an empty string is a title with empty text, not a removal).
 */
export function setChartTitle(space: ChartSpace, text: string | undefined): ChartSpace {
	return editChartSpace(space, (draft) => {
		if (text === undefined) {
			delete draft.title;
			draft.autoTitleDeleted = true;
			return;
		}
		draft.title = chartTitleWithText(draft.title, text);
		draft.autoTitleDeleted = false;
	});
}

/**
 * Shows or hides the chart title (Chart Elements > Chart Title). Showing a chart that has none adds
 * an automatic title (no text: Office shows the series name of a one-series chart, else "Chart
 * Title"); hiding removes it and sets `c:autoTitleDeleted`. Semantics: PowerPoint and Excel.
 */
export function setChartTitleVisible(space: ChartSpace, visible: boolean): ChartSpace {
	return editChartSpace(space, (draft) => {
		if (!visible) {
			delete draft.title;
			draft.autoTitleDeleted = true;
			return;
		}
		draft.title ??= { overlay: false };
		draft.autoTitleDeleted = false;
	});
}

/** Whether the title is drawn over the plot area (`c:overlay`); no effect without a title. */
export function setChartTitleOverlay(space: ChartSpace, overlay: boolean): ChartSpace {
	return editChartSpace(space, (draft) => {
		if (draft.title) draft.title.overlay = overlay;
	});
}
