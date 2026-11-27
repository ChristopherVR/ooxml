// The number format a PivotTable value field names (`dataField/@numFmtId`): a built-in format or
// one of the workbook's own `numFmts` (ECMA-376 Part 1, 18.8.30, 18.8.31).
import { RELATIONSHIP_TYPES } from '../../opc/index';
import { NS, children, first, parseXml } from '../../xml/index';
import { builtinFormat } from '../read/builtin-formats';
import type { SourceIndex } from '../read/package';
import { att } from '../read/xml-util';

const formats = new WeakMap<SourceIndex, ReadonlyMap<number, string>>();

/** The custom number formats of the package's styles part, by id. */
function customFormats(source: SourceIndex): ReadonlyMap<number, string> {
	const known = formats.get(source);
	if (known) return known;
	const book = source.workbookPart();
	const part = book ? source.targetOfType(book, RELATIONSHIP_TYPES.styles) : undefined;
	const xml = part ? source.text(part) : undefined;
	const map = new Map<number, string>();
	if (xml) {
		const root = parseXml(xml).documentElement;
		for (const node of children(first(root, 'numFmts', NS.x) ?? root, 'numFmt', NS.x)) {
			const id = Number(att(node, 'numFmtId'));
			const code = att(node, 'formatCode');
			if (Number.isInteger(id) && code !== undefined) map.set(id, code);
		}
	}
	formats.set(source, map);
	return map;
}

/**
 * The format code of a value field's `numFmtId`, undefined for General (0) or no attribute. An id
 * the package does not define is refused, so the report is kept as it was written.
 */
export function pivotNumberFormat(
	source: SourceIndex,
	value: string | undefined,
): string | undefined {
	if (value === undefined) return undefined;
	const id = Number(value);
	if (!Number.isInteger(id) || id < 0) throw new Error('Unsupported value field number format');
	if (id === 0) return undefined;
	const code = customFormats(source).get(id) ?? builtinFormat(id);
	if (code === undefined) throw new Error('Unsupported value field number format');
	return code === 'General' ? undefined : code;
}
