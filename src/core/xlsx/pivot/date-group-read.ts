import { NS, children, elements, first, type XmlElement } from '../../xml/index';
import type { CellValue, PivotDateGroup, PivotDateUnit } from '../model';
import { att, decodeEscapes } from '../read/xml-util';
import {
	PIVOT_DATE_UNITS,
	dateGroupDays,
	dateGroupFieldName,
	dateGroupRange,
	dateUnitLabels,
} from './date-groups';
import { isoToSerial } from './dates';

const unsupported = (message: string) => new Error(`Unsupported pivot date group: ${message}`);
const close = (a: number, b: number) => Math.abs(a - b) < 1e-6;

/** The cache fields that exist only to hold a date group: they follow every database field. */
export const databaseFieldCount = (fields: XmlElement[]): number => {
	const count = fields.findIndex((field) => att(field, 'databaseField') === '0');
	const first = count < 0 ? fields.length : count;
	if (fields.slice(first).some((field) => att(field, 'databaseField') !== '0'))
		throw unsupported('database fields follow group fields');
	return first;
};

/**
 * Date groups: a date field with `fieldGroup/@par` plus one trailing cache field per unit
 * (`rangePr/@groupBy`) whose group items are the ones the model generates. Anything else, such as
 * discrete groups or a changed item list, is preserved unsupported.
 */
export function readDateGroups(
	fields: XmlElement[],
	count: number,
	names: string[],
	dateFields: readonly string[],
	records: readonly CellValue[][],
	date1904: boolean,
): PivotDateGroup[] {
	const derived = fields.slice(count);
	const bases = fields
		.slice(0, count)
		.map((field, index) => [first(field, 'fieldGroup', NS.x), index] as const)
		.filter(([group]) => group && att(group, 'par') !== undefined);
	if (!derived.length) {
		if (bases.length) throw unsupported('a group parent without group fields');
		return [];
	}
	const groups: PivotDateGroup[] = [];
	const members = new Map<number, { index: number; unit: PivotDateUnit; node: XmlElement }[]>();
	let last = -1;
	derived.forEach((node, offset) => {
		const group = first(node, 'fieldGroup', NS.x);
		const range = first(group, 'rangePr', NS.x);
		const items = first(group, 'groupItems', NS.x);
		const base = Number(att(group, 'base'));
		const unit = att(range, 'groupBy') as PivotDateUnit | undefined;
		if (
			!group ||
			!range ||
			!items ||
			!Number.isInteger(base) ||
			base < 0 ||
			base >= count ||
			!unit ||
			!PIVOT_DATE_UNITS.includes(unit) ||
			att(group, 'par') !== undefined ||
			att(node, 'numFmtId') !== '0' ||
			elements(node).some((child) => child.localName !== 'fieldGroup') ||
			elements(group).some((child) => !['rangePr', 'groupItems'].includes(child.localName ?? ''))
		)
			throw unsupported('group field');
		if (base !== last && members.has(base)) throw unsupported('interleaved group fields');
		last = base;
		const list = members.get(base) ?? [];
		list.push({ index: count + offset, unit, node });
		members.set(base, list);
	});
	if (bases.some(([, index]) => !members.has(index))) throw unsupported('a group parent');
	for (const [base, list] of members) {
		const field = names[base] ?? '';
		const parent = first(fields[base], 'fieldGroup', NS.x);
		const units = list.map((entry) => entry.unit);
		if (
			!dateFields.includes(field) ||
			Number(att(parent, 'par')) !== list.at(-1)!.index ||
			units.some(
				(unit, i) =>
					i > 0 && PIVOT_DATE_UNITS.indexOf(unit) <= PIVOT_DATE_UNITS.indexOf(units[i - 1]!),
			)
		)
			throw unsupported('group parent');
		const ranges = list.map((entry) =>
			first(first(entry.node, 'fieldGroup', NS.x), 'rangePr', NS.x)!,
		);
		const sample = ranges[0]!;
		const autoStart = att(sample, 'autoStart') !== '0';
		const autoEnd = att(sample, 'autoEnd') !== '0';
		const start = isoToSerial(att(sample, 'startDate') ?? '', date1904);
		const end = isoToSerial(att(sample, 'endDate') ?? '', date1904);
		if (
			start === undefined ||
			end === undefined ||
			ranges.some(
				(range) =>
					att(range, 'startDate') !== att(sample, 'startDate') ||
					att(range, 'endDate') !== att(sample, 'endDate') ||
					(att(range, 'autoStart') !== '0') !== autoStart ||
					(att(range, 'autoEnd') !== '0') !== autoEnd,
			)
		)
			throw unsupported('range');
		const interval = att(sample, 'groupInterval');
		const days = interval === undefined ? 1 : Number(interval);
		if (
			!Number.isInteger(days) ||
			days < 1 ||
			(interval !== undefined && units.join() !== 'days') ||
			ranges.some((range, i) => i > 0 && att(range, 'groupInterval') !== undefined)
		)
			throw unsupported('interval');
		const group: PivotDateGroup = {
			field,
			by: units,
			...(autoStart ? {} : { start }),
			...(autoEnd ? {} : { end }),
			...(days > 1 ? { days } : {}),
		};
		const serials = records.map((record) => record[base]);
		if (serials.some((value) => typeof value !== 'number')) throw unsupported('non-date value');
		const expected = dateGroupRange(group, serials as number[]);
		if (!close(expected.start, start) || !close(expected.end, end))
			throw unsupported('stale range');
		for (const entry of list) {
			const labels = dateUnitLabels(entry.unit, expected, dateGroupDays(group), date1904);
			const node = first(first(entry.node, 'fieldGroup', NS.x), 'groupItems', NS.x)!;
			const actual = children(node, 's', NS.x).map((item) => decodeEscapes(att(item, 'v') ?? ''));
			if (
				att(entry.node, 'name') === undefined ||
				decodeEscapes(att(entry.node, 'name')!) !== dateGroupFieldName(entry.unit, field) ||
				elements(node).length !== actual.length ||
				actual.length !== labels.length ||
				actual.some((label, i) => label !== labels[i])
			)
				throw unsupported('group items');
		}
		groups.push(group);
	}
	return groups;
}
