/** Native DATE caches are wall-clock ISO text, without a timezone. */
export function visioIsoDateSerial(text: string): number | undefined {
	const match = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?)?$/.exec(
		text.trim(),
	);
	if (!match) return undefined;
	const year = Number(match[1]),
		month = Number(match[2]),
		day = Number(match[3]);
	const hour = Number(match[4] ?? 0),
		minute = Number(match[5] ?? 0),
		second = Number(match[6] ?? 0);
	const millisecond = Number((match[7] ?? '').padEnd(3, '0').slice(0, 3));
	if (year < 1 || hour > 23 || minute > 59 || second > 59) return undefined;
	const date = new Date(0);
	date.setUTCFullYear(year, month - 1, day);
	date.setUTCHours(hour, minute, second, millisecond);
	if (
		date.getUTCFullYear() !== year ||
		date.getUTCMonth() !== month - 1 ||
		date.getUTCDate() !== day
	)
		return undefined;
	return (date.getTime() - Date.UTC(1899, 11, 30)) / 86_400_000;
}
