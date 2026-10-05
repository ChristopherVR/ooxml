import { dateToSerial } from '../numfmt/date.js';

const DAY_MS = 86_400_000;
const EPOCH_1900 = Date.UTC(1899, 11, 30);

/**
 * The serial number of an ISO 8601 date (`t="d"` cells), read as wall time. Serials after
 * 1900-02-28 follow Excel's 1900 leap-year convention.
 */
export function isoToSerial(iso: string, date1904: boolean): number | undefined {
	const match =
		/^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?)?(?:Z|[+-]\d{2}:?\d{2})?$/.exec(
			iso.trim(),
		);
	if (!match) {
		const time = /^(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?$/.exec(iso.trim());
		if (!time) return undefined;
		const seconds = Number(time[1]) * 3600 + Number(time[2]) * 60 + Number(time[3] ?? 0);
		return (seconds + Number(`0.${time[4] ?? '0'}`)) / 86_400;
	}
	const ms = Date.UTC(
		Number(match[1]),
		Number(match[2]) - 1,
		Number(match[3]),
		Number(match[4] ?? 0),
		Number(match[5] ?? 0),
		Number(match[6] ?? 0),
		Math.round(Number(`0.${match[7] ?? '0'}`) * 1000),
	);
	// Times stored on 1899-12-30 (serial 0 plus a fraction) are not shifted by the 1900 leap bug.
	if (!date1904 && ms < EPOCH_1900 + DAY_MS) return (ms - EPOCH_1900) / DAY_MS;
	return dateToSerial(new Date(ms), date1904);
}
