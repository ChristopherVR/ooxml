// worksheet/ignoredErrors: the Error Checking rules switched off for ranges (Ignore Error).
import type { XmlElement } from '../../xml/index';
import { parseRange } from '../address';
import { IGNORED_ERROR_FLAGS, type IgnoredError } from '../model';
import { att, xChildren } from './xml-util';

export function readIgnoredErrors(node: XmlElement | undefined): IgnoredError[] {
	if (!node) return [];
	const out: IgnoredError[] = [];
	for (const child of xChildren(node, 'ignoredError')) {
		const ranges = (att(child, 'sqref') ?? '')
			.split(/\s+/)
			.map((ref) => parseRange(ref))
			.filter((range) => range !== undefined);
		const flags = IGNORED_ERROR_FLAGS.filter((flag) => {
			const value = att(child, flag);
			return value === '1' || value === 'true';
		});
		if (ranges.length && flags.length) out.push({ ranges, flags });
	}
	return out;
}
