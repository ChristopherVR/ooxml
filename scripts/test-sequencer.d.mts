export function balanceTestFiles(
	files: string[],
	timings: Record<string, number>,
	count: number,
): Array<{ files: string[]; duration: number }>;

export function durationSequencer<Context, Sequencer>(
	BaseSequencer: new (context: Context) => Sequencer,
	timingFile: URL,
): new (context: Context) => Sequencer;
