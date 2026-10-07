import type { InlineImage, TextRun } from './model';

/** A media relationship identifies the source drawing independently of preceding text runs. */
export function samePictureSource(a: InlineImage | undefined, b: InlineImage | undefined): boolean {
	return !!a?.partName && a.partName === b?.partName && a.relId === b?.relId;
}

export function inlineSourceLookup(
	runs: TextRun[],
	base: TextRun[] | undefined,
): (index: number) => number | undefined {
	const usedPictures = new Set<number>();
	return (index) => {
		const run = runs[index]!;
		if (run.equation)
			return base?.findIndex(
				(source) =>
					source.equation?.omml === run.equation?.omml &&
					source.equation?.display === run.equation?.display,
			);
		if (!run.image?.partName) return index;
		const candidates = (base ?? []).flatMap((source, at) =>
			samePictureSource(source.image, run.image) ? [at] : [],
		);
		// Multiple drawings can share media but carry different unmodeled DrawingML. Without a
		// stable drawing identity, removing one must not attach another drawing's XML to its peer.
		if (
			candidates.length > 1 &&
			runs.filter((current) => samePictureSource(current.image, run.image)).length !==
				candidates.length
		)
			throw new Error(
				'Cannot safely match pictures sharing the same media after their count changes. The original DOCX package remains unchanged.',
			);
		const source = candidates.find((at) => !usedPictures.has(at));
		if (source !== undefined) usedPictures.add(source);
		return source;
	};
}
