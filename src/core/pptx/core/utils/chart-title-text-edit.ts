import type { PptxChartTitleRun, XmlObject } from '../types';
import { findKey, type GetLocalName } from './chart-title-xml-ops';

/** Patch a single rich paragraph without rebuilding its run or paragraph properties. */
export function patchChartTitleRunText(
	title: XmlObject,
	runs: readonly PptxChartTitleRun[],
	local: GetLocalName,
): boolean {
	const tx = title[findKey(title, 'tx', local) ?? ''] as XmlObject | undefined;
	const rich = tx?.[findKey(tx, 'rich', local) ?? ''] as XmlObject | undefined;
	const raw = rich?.[findKey(rich, 'p', local) ?? ''];
	const paragraphs = Array.isArray(raw) ? raw : raw ? [raw] : [];
	if (paragraphs.length !== 1) return false;
	const paragraph = paragraphs[0] as XmlObject;
	if (findKey(paragraph, 'br', local) || findKey(paragraph, 'fld', local)) return false;
	const rawRuns = paragraph[findKey(paragraph, 'r', local) ?? ''];
	const original = Array.isArray(rawRuns) ? rawRuns : rawRuns ? [rawRuns] : [];
	if (original.length !== runs.length) return false;
	const keys = original.map((run) => findKey(run as XmlObject, 't', local));
	if (keys.some((key) => !key)) return false;
	for (const [index, value] of original.entries()) {
		const run = value as XmlObject,
			key = keys[index]!;
		const text = run[key];
		run[key] =
			text && typeof text === 'object' && !Array.isArray(text)
				? { ...(text as XmlObject), '#text': runs[index]!.text }
				: runs[index]!.text;
	}
	return true;
}
