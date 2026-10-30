import { formatRange } from '../address';
import type { IgnoredError } from '../model';

export function ignoredErrorsXml(list: readonly IgnoredError[] | undefined): string {
	const items = (list ?? []).filter((e) => e.ranges.length && e.flags.length);
	if (!items.length) return '';
	const body = items
		.map((e) => {
			const flags = e.flags.map((flag) => ` ${flag}="1"`).join('');
			return `<ignoredError sqref="${e.ranges.map(formatRange).join(' ')}"${flags}/>`;
		})
		.join('');
	return `<ignoredErrors>${body}</ignoredErrors>`;
}
