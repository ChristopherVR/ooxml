export function balanceFiles(
	files: string[],
	totals: Record<string, number>,
	shards: number,
): Array<{ files: string[]; duration: number }>;
export function fileDurations(history: unknown, project: string): Record<string, number>;
