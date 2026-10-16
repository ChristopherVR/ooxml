import { NS, children, elements, first, type XmlElement } from '../../xml/index';
import type { PivotFieldGroup } from '../model';
import { att, decodeEscapes } from '../read/xml-util';
import { numberGroupLabels } from './groups';

/**
 * Number groups of the cache fields: a `fieldGroup` on the field itself with an integer
 * `rangePr` (start, end, interval) whose item labels are the ones the model generates. Date
 * groups (extra cache fields, `groupBy`) and discrete groups are preserved unsupported.
 */
export function readFieldGroups(fields: XmlElement[], names: string[]): PivotFieldGroup[] {
	const groups: PivotFieldGroup[] = [];
	fields.forEach((field, index) => {
		const group = first(field, 'fieldGroup', NS.x);
		// A date field's `par` link belongs to its date groups (read separately).
		if (!group || (att(group, 'par') !== undefined && !first(group, 'rangePr', NS.x))) return;
		const range = first(group, 'rangePr', NS.x);
		const list = first(group, 'groupItems', NS.x);
		const number = (name: string) => {
			const raw = att(range, name);
			return raw !== undefined && /^-?\d+$/.test(raw) ? Number(raw) : undefined;
		};
		const start = number('startNum');
		const end = number('endNum');
		const by = number('groupInterval');
		if (
			field.hasAttribute('databaseField') ||
			att(group, 'base') !== String(index) ||
			att(group, 'par') !== undefined ||
			!range ||
			!list ||
			start === undefined ||
			end === undefined ||
			by === undefined ||
			by < 1 ||
			start >= end ||
			att(range, 'groupBy') !== undefined ||
			elements(group).some((node) => !['rangePr', 'groupItems'].includes(node.localName ?? ''))
		)
			throw new Error('Unsupported pivot field group');
		const labels = children(list, 's', NS.x).map((node) => decodeEscapes(att(node, 'v') ?? ''));
		const expected = numberGroupLabels({ start, end, by });
		if (
			elements(list).length !== labels.length ||
			labels.length !== expected.length ||
			labels.some((label, i) => label !== expected[i])
		)
			throw new Error('Unsupported pivot group items');
		groups.push({ field: names[index] ?? '', start, end, by });
	});
	return groups;
}
