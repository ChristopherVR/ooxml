/**
 * Reassign edited flat text to its existing styled runs using surviving run boundaries.
 * Unchanged runs keep their text; inserted trailing text belongs to the last run. Returns
 * undefined when no later boundary anchors a replacement, so callers can choose a full rewrite.
 * Extracted from PowerPoint chart title editing; no document model is required.
 */
export function distributeRunsText(
	oldTexts: readonly string[],
	newText: string,
): string[] | undefined {
	if (!oldTexts.length) return undefined;
	const result: string[] = new Array(oldTexts.length).fill('');
	let pos = 0;
	let index = 0;
	while (index < oldTexts.length) {
		const text = oldTexts[index] ?? '';
		if (newText.startsWith(text, pos)) {
			result[index] = text;
			pos += text.length;
			index++;
			continue;
		}
		let resumeAt = -1;
		let anchorIndex = -1;
		for (let candidate = index + 1; candidate < oldTexts.length; candidate++) {
			const candidateText = oldTexts[candidate];
			if (!candidateText) continue;
			const found = newText.indexOf(candidateText, pos);
			if (found !== -1) {
				resumeAt = found;
				anchorIndex = candidate;
				break;
			}
		}
		if (anchorIndex === -1) return undefined;
		result[index] = newText.slice(pos, resumeAt);
		pos = resumeAt;
		index = anchorIndex;
	}
	if (pos < newText.length) result[result.length - 1] += newText.slice(pos);
	return result;
}

/** Preserve runs when boundaries survive; otherwise retain the first run's style only. */
export function realignOrCollapseRuns<T extends { text: string }>(
	runs: readonly T[],
	text: string,
): T[] {
	const first = runs[0];
	if (!first) return [];
	const distributed = distributeRunsText(
		runs.map((run) => run.text),
		text,
	);
	return distributed
		? runs.map((run, index) => ({ ...run, text: distributed[index] ?? '' }))
		: [{ ...first, text }];
}
