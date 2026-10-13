import { visioIsoDateSerial } from './date-cache';
import { formatVisioFieldValue } from './text-fields';

/** Native system pictures are locale-dependent. Admit only a valid stored ISO profile,
 * with the measured English twelve-hour suffix for pictures 211/212. */
export function systemDatePicture(date: Date, code: string, cached: string): string | undefined {
	if (
		!Number.isFinite(date.getTime()) ||
		date.getTime() < Date.UTC(1899, 11, 30) ||
		date.getUTCFullYear() > 9999
	)
		return undefined;
	if (code === '204') return formatVisioFieldValue({ kind: 'date', date }, '{{yyyy-MM-dd}}');
	if (code === '20' || code === '200') {
		if (!/^\d{4}-\d{2}-\d{2}$/.test(cached) || visioIsoDateSerial(cached) === undefined)
			return undefined;
		return formatVisioFieldValue({ kind: 'date', date }, '{{yyyy-MM-dd}}');
	}
	if (code !== '211' && code !== '212') return undefined;
	const match = /^(\d{4}-\d{2}-\d{2}) ([1-9]|1[0-2]):([0-5]\d)(?::([0-5]\d))? (AM|PM)$/.exec(
		cached,
	);
	if (!match || (code === '212') !== (match[4] !== undefined)) return undefined;
	const hour = (Number(match[2]) % 12) + (match[5] === 'PM' ? 12 : 0);
	if (
		visioIsoDateSerial(
			`${match[1]}T${String(hour).padStart(2, '0')}:${match[3]}:${match[4] ?? '00'}`,
		) === undefined
	)
		return undefined;
	return formatVisioFieldValue(
		{ kind: 'date', date },
		code === '211' ? '{{yyyy-MM-dd h:mm tt}}' : '{{yyyy-MM-dd h:mm:ss tt}}',
	);
}
