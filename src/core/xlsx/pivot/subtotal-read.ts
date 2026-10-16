import { NS, children, first, type XmlElement } from '../../xml/index';
import type { PivotFieldSubtotals } from '../model';
import { att } from '../read/xml-util';
import { PIVOT_SUBTOTALS, PIVOT_SUBTOTAL_ORDER } from './subtotals';

/**
 * The subtotal settings of row fields that differ from Automatic: `defaultSubtotal="0"` is none
 * (an empty list) and the `sumSubtotal`-style flags are custom functions. Excel leaves
 * `defaultSubtotal` out when it writes custom flags, so a stated `1` next to flags is ambiguous.
 */
export function fieldSubtotalsOf(
	definitions: readonly XmlElement[],
	names: readonly (string | undefined)[],
): PivotFieldSubtotals[] {
	const settings: PivotFieldSubtotals[] = [];
	definitions.forEach((definition, index) => {
		const field = names[index];
		if (!field) throw new Error('Invalid pivot row field');
		const functions = PIVOT_SUBTOTAL_ORDER.filter(
			(fn) => att(definition, PIVOT_SUBTOTALS[fn].attr) === '1',
		);
		const automatic = !functions.length && att(definition, 'defaultSubtotal') !== '0';
		if (functions.length && att(definition, 'defaultSubtotal') === '1')
			throw new Error('Automatic and custom subtotals together are preserved');
		// The subtotal items that close the item list must be the ones the flags describe.
		const listed = children(first(definition, 'items', NS.x) ?? definition, 'item', NS.x)
			.map((item) => att(item, 't'))
			.filter((type) => type !== undefined && type !== 'data');
		const expected = automatic ? ['default'] : functions.map((fn) => PIVOT_SUBTOTALS[fn].item);
		if (listed.join() !== expected.join()) throw new Error('Unsupported pivot subtotal items');
		if (!automatic) settings.push({ field, functions });
	});
	return settings;
}
