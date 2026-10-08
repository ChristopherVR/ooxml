import { attribute, children } from './sheet';
import { fail } from './package-common';
import { visioFormulaCachedValue } from './formula';

const delegatedSources = new WeakMap<Element, Element | undefined>();

/** Delegated style caches are admitted only when they agree with a proven ancestor. */
export function delegatedFormattingCell(cell: Element | undefined): Element | undefined {
	const seen = new Set<Element>();
	while (cell && attribute(cell, 'F') === 'Inh') {
		const ancestor = delegatedSources.get(cell);
		if (!ancestor || seen.has(cell) || seen.size >= 64 || cell.hasAttribute('E'))
			fail('EDIT_PROTECTED_CELL', 'Inherited formatting has no proven source.');
		seen.add(cell);
		const current = attribute(cell, 'V'),
			previous = attribute(ancestor, 'V');
		if (current !== previous || attribute(cell, 'U') !== attribute(ancestor, 'U')) {
			let agrees = false;
			try {
				const a = visioFormulaCachedValue(current ?? '', attribute(cell, 'U'));
				const b = visioFormulaCachedValue(previous ?? '', attribute(ancestor, 'U'));
				agrees = a.value === b.value && a.unit === b.unit;
			} catch {
				/* Nonnumeric caches must agree exactly. */
			}
			if (!agrees)
				fail(
					'EDIT_PROTECTED_CELL',
					'Inherited formatting cache is stale or has incompatible units.',
				);
		}
		cell = ancestor;
	}
	return cell;
}

/** Canonical XML indices prevent aliases such as 01 from creating duplicate rows. */
export function formattingIndex(value: string | undefined): string {
	if (value === undefined || !/^(0|[1-9]\d{0,9})$/.test(value) || Number(value) > 0xffffffff)
		fail('UNSUPPORTED_FORMAT_EDIT', 'Formatting requires canonical unsigned row indices.');
	return value;
}
export function formattingRows(sheet: Element, sectionName: string): Map<string, Element> {
	const sections = children(sheet, 'Section').filter((section) => {
		const name = attribute(section, 'N');
		if (name?.toLowerCase() === sectionName.toLowerCase() && name !== sectionName)
			fail('UNSUPPORTED_FORMAT_EDIT', 'Noncanonical formatting section names are unsupported.');
		return name === sectionName;
	});
	if (sections.length > 1) fail('UNSUPPORTED_FORMAT_EDIT', 'Formatting section is ambiguous.');
	const section = sections[0];
	const result = new Map<string, Element>();
	if (!section) return result;
	if (section.hasAttribute('Del') || (attribute(section, 'IX') ?? '0') !== '0')
		fail('UNSUPPORTED_FORMAT_EDIT', 'Formatting requires one live section with index zero.');
	for (const row of children(section, 'Row')) {
		const index = formattingIndex(attribute(row, 'IX') ?? '0');
		if (
			row.hasAttribute('Del') ||
			row.hasAttribute('N') ||
			row.hasAttribute('T') ||
			result.has(index)
		)
			fail('UNSUPPORTED_FORMAT_EDIT', 'Formatting requires unique live indexed rows.');
		result.set(index, row);
		if (result.size > 10_000) fail('UNSUPPORTED_FORMAT_EDIT', 'Formatting row limit exceeded.');
	}
	return result;
}
export function formattingRow(
	sheet: Element,
	sectionName: string,
	index = '0',
): Element | undefined {
	return formattingRows(sheet, sectionName).get(index);
}
export type FormattingRowSources = ReadonlyMap<string, ReadonlyMap<string, Element>>;

/** Mirror mergeSheets while retaining source cell attributes for protection and dependency proof.
 * Each local row inherits the previous sheet's matching row or row zero, never a local sibling.
 */
export function mergeFormattingRows(
	base: FormattingRowSources,
	sheet: Element,
	sectionName: string,
	cells: (row: Element) => ReadonlyMap<string, Element>,
): FormattingRowSources {
	const result = new Map(base);
	for (const [index, row] of formattingRows(sheet, sectionName)) {
		const merged = new Map(base.get(index) ?? base.get('0'));
		for (const [name, cell] of cells(row)) {
			// A cache-bearing Inh cell remains the parser's effective saved value.
			// Admission rejects overriding that unresolved source rather than trusting its cache.
			if (attribute(cell, 'F') === 'Inh' && !cell.hasAttribute('V') && !cell.hasAttribute('E'))
				continue;
			const ancestor = merged.get(name);
			const inheritedUnit = attribute(ancestor, 'U');
			let effective = cell;
			if (!cell.hasAttribute('U') && inheritedUnit !== undefined) {
				effective = cell.cloneNode(true) as Element;
				effective.setAttribute('U', inheritedUnit);
			}
			if (attribute(cell, 'F') === 'Inh') delegatedSources.set(effective, ancestor);
			merged.set(name, effective);
		}
		result.set(index, merged);
		if (result.size > 10_000) fail('UNSUPPORTED_FORMAT_EDIT', 'Formatting row limit exceeded.');
	}
	return result;
}

/** Text and field nodes are retained; only proven indexed formatting markers are admitted. */
export function assertFormattingText(
	shape: Element,
	characters?: FormattingRowSources,
	paragraphs?: FormattingRowSources,
): void {
	const texts = children(shape, 'Text');
	if (texts.length !== 1)
		fail('UNSUPPORTED_FORMAT_EDIT', 'Text formatting requires one existing local Text.');
	let fields: Map<string, Element> | undefined;
	for (const part of Array.from(texts[0]!.childNodes)) {
		if (part.nodeType === 3 || part.nodeType === 4) continue;
		const marker = part as Element;
		if (
			part.nodeType !== 1 ||
			marker.namespaceURI !== shape.namespaceURI ||
			!['cp', 'pp', 'tp', 'fld'].includes(marker.localName)
		)
			fail('UNSUPPORTED_FORMAT_EDIT', 'Unknown text markup cannot be formatted.');
		const index = formattingIndex(attribute(marker, 'IX'));
		if (marker.localName === 'fld') {
			fields ??= formattingRows(shape, 'Field');
			if (
				!fields.has(index) ||
				Array.from(marker.childNodes).some((node) => ![3, 4].includes(node.nodeType))
			)
				fail(
					'UNSUPPORTED_FORMAT_EDIT',
					'Text fields require a live local field row and plain cached text.',
				);
			continue;
		}
		const rows =
			marker.localName === 'cp' ? characters : marker.localName === 'pp' ? paragraphs : undefined;
		if (marker.childNodes.length || (index !== '0' && rows && !rows.has(index)))
			fail(
				'UNSUPPORTED_FORMAT_EDIT',
				'Formatting markers require existing rows and no nested markup.',
			);
	}
}
