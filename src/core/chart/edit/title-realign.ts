import { distributeRunsText } from '../../text/realign-runs';
import { NS, parseXml } from '../../xml/index';
import type { ChartTextBody } from '../model-series';
import { chartTextBodyFromXml } from '../write-shape';
import { innerXml } from '../xml-fragment';

/** Keep the original runs and their opaque properties when a local text edit retains boundaries. */
export function realignChartTitleBody(
	body: ChartTextBody,
	text: string,
): ChartTextBody | undefined {
	if (body.paragraphs.length !== 1) return undefined;
	const runs = body.paragraphs[0]!.runs;
	const distributed = distributeRunsText(
		runs.map((run) => run.text),
		text,
	);
	if (!distributed) return undefined;
	if (body.sourceXml !== undefined) {
		const root = parseXml(
			`<c:rich xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}">${body.sourceXml}</c:rich>`,
		).documentElement;
		if (
			root.getElementsByTagNameNS(NS.a, 'fld').length ||
			root.getElementsByTagNameNS(NS.a, 'br').length
		)
			return undefined;
		const original = [...root.getElementsByTagNameNS(NS.a, 'r')];
		if (original.length !== distributed.length) return undefined;
		const texts = original.map((run) => run.getElementsByTagNameNS(NS.a, 't')[0]);
		if (texts.some((node) => !node)) return undefined;
		texts.forEach((node, index) => {
			node!.textContent = distributed[index]!;
		});
		return chartTextBodyFromXml(innerXml(root));
	}
	return {
		...body,
		paragraphs: [
			{
				...body.paragraphs[0]!,
				runs: runs.map((run, index) => ({ ...run, text: distributed[index]! })),
			},
		],
		text,
	};
}
