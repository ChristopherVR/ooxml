import type { NumberingCatalog } from './numbering-model';
import { resolveNumberingLevel } from './numbering-parse';

export function assertListStartValue(level: number, start: number): void {
	if (!Number.isInteger(level) || level < 0 || level > 8)
		throw new Error('List level must be an integer from 0 through 8.');
	// ST_DecimalNumber is xsd:integer; Word's Open XML SDK maps its value to Int32Value.
	// Bound this authoring action to that interoperable range, without narrowing parsed input.
	if (!Number.isInteger(start) || start < 0 || start > 2147483647)
		throw new Error('List start must be an integer from 0 through 2147483647.');
}

/**
 * Changes one existing instance's start without modifying its shared abstract definition.
 * Full level overrides are unsupported here: they require a separate, preserving edit action.
 */
export function setListStartOverride(
	catalog: NumberingCatalog,
	numId: number,
	level: number,
	start: number,
): NumberingCatalog {
	assertListStartValue(level, start);
	if (!Number.isSafeInteger(numId) || numId <= 0) throw new Error('Invalid list instance id.');
	const id = String(numId);
	const num = catalog.nums[id];
	if (!num || !resolveNumberingLevel(catalog, id, level))
		throw new Error('List instance or level does not exist.');
	const override = num.levelOverrides?.[level];
	if (override?.lvl) throw new Error('Changing starts with a full level override is unsupported.');
	if (override?.startOverride === start) return catalog;
	return {
		...catalog,
		nums: {
			...catalog.nums,
			[id]: {
				...num,
				levelOverrides: {
					...num.levelOverrides,
					[level]: { ...override, startOverride: start },
				},
			},
		},
	};
}
