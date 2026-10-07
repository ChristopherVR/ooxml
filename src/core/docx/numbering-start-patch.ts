import type { NumberingCatalog, NumDefinition, NumLevelOverride } from './numbering-model';
import { assertListStartValue } from './numbering-start-edit';
import { resolveNumberingLevel } from './numbering-parse';
import { children, getW, makeW, WORD_NS, type XmlDocument, type XmlElement } from './xml';

interface StartChange {
	numId: string;
	level: number;
	start: number;
}

function withoutStarts(num: NumDefinition): object {
	const { levelOverrides, ...properties } = num;
	const kept: Record<string, Omit<NumLevelOverride, 'startOverride'>> = {};
	for (const [level, override] of Object.entries(levelOverrides ?? {})) {
		const { startOverride: _start, ...rest } = override;
		if (Object.keys(rest).length) kept[level] = rest;
	}
	return { ...properties, ...(Object.keys(kept).length ? { levelOverrides: kept } : {}) };
}

/** Validate every existing-instance delta before touching package XML. */
export function numberingStartChanges(
	prior: NumberingCatalog | undefined,
	next: NumberingCatalog,
): StartChange[] {
	const changes: StartChange[] = [];
	for (const [id, previous] of Object.entries(prior?.nums ?? {})) {
		const current = next.nums[id];
		if (
			!current ||
			JSON.stringify(withoutStarts(current)) !== JSON.stringify(withoutStarts(previous))
		)
			throw new Error(
				`Cannot edit numbering definition num "${id}" except its supported start overrides.`,
			);
		for (const key of new Set([
			...Object.keys(previous.levelOverrides ?? {}),
			...Object.keys(current.levelOverrides ?? {}),
		])) {
			const level = Number(key);
			const old = previous.levelOverrides?.[level];
			const value = current.levelOverrides?.[level];
			if (old?.startOverride === value?.startOverride) continue;
			if (value?.startOverride === undefined)
				throw new Error('Removing a list start override is unsupported.');
			assertListStartValue(level, value.startOverride);
			if (String(level) !== key || !prior || !resolveNumberingLevel(prior, id, level))
				throw new Error('List instance or level does not exist.');
			if (old?.lvl || value.lvl)
				throw new Error('Changing starts with a full level override is unsupported.');
			changes.push({ numId: id, level, start: value.startOverride });
		}
	}
	return changes;
}

/** Preflight source identities, then update only startOverride values in the original XML tree. */
export function patchNumberingStarts(doc: XmlDocument, changes: readonly StartChange[]): void {
	const sameNumber = (source: string | undefined, expected: string) =>
		source !== undefined && /^[+-]?\d+$/.test(source) && Number(source) === Number(expected);
	const prepared = changes.map((change) => {
		const nums = children(doc.documentElement, 'num').filter((num) =>
			sameNumber(getW(num, 'numId'), change.numId),
		);
		if (nums.length !== 1 || getW(nums[0], 'numId') !== change.numId)
			throw new Error('Cannot patch an ambiguous or missing source list instance.');
		const num = nums[0]!;
		const overrides = children(num, 'lvlOverride').filter((node) =>
			sameNumber(getW(node, 'ilvl'), String(change.level)),
		);
		if (overrides.length > 1) throw new Error('Cannot patch duplicate source level overrides.');
		const override = overrides[0];
		if (override && getW(override, 'ilvl') !== String(change.level))
			throw new Error('Cannot patch a noncanonical source level override.');
		if (
			override &&
			(children(override, 'lvl').length || children(override, 'startOverride').length > 1)
		)
			throw new Error('Cannot patch a full or ambiguous source level override.');
		return { change, num, override };
	});
	for (const { change, num, override: existing } of prepared) {
		const override: XmlElement = existing ?? makeW(doc, 'lvlOverride');
		if (!existing) {
			override.setAttributeNS(WORD_NS, 'w:ilvl', String(change.level));
			const after = children(num, 'lvlOverride').at(-1) ?? children(num, 'abstractNumId')[0];
			num.insertBefore(override, after?.nextSibling ?? null);
		}
		const start = children(override, 'startOverride')[0] ?? makeW(doc, 'startOverride');
		start.setAttributeNS(WORD_NS, 'w:val', String(change.start));
		if (!start.parentNode) override.insertBefore(start, override.firstChild);
	}
}
