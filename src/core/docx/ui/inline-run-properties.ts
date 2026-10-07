import type { AttributeSpec, Node } from 'prosemirror-model';
import type { TextRun } from '../model';
import { DIRECT_RUN_PROPERTY_KEYS } from '../run-formatting';

const attributeName = (key: string) => `runFormat_${key}`;
const directKeys = new Set<string>(DIRECT_RUN_PROPERTY_KEYS);
export const RUN_PROPERTY_ATTRIBUTES = new Set(DIRECT_RUN_PROPERTY_KEYS.map(attributeName));
/** Null means no override; the string "null" removes a property from the imported basis. */
export const inlineRunFormattingAttrs: Record<string, AttributeSpec> = Object.fromEntries(
	[...RUN_PROPERTY_ATTRIBUTES].map((name) => [name, { default: null }]),
);

export function inlineRunProperties(node: Node): Partial<TextRun> {
	const properties: Record<string, unknown> =
		typeof node.attrs.format === 'string' ? JSON.parse(node.attrs.format) : {};
	for (const key of DIRECT_RUN_PROPERTY_KEYS) {
		const override = node.attrs[attributeName(key)];
		if (typeof override !== 'string') continue;
		const value: unknown = JSON.parse(override);
		if (value === null) delete properties[key];
		else properties[key] = value;
	}
	return properties as Partial<TextRun>;
}

/** Copies serialize effective properties, so they do not depend on the source room's overrides. */
export function inlineRunPropertiesDomAttrs(node: Node): Record<string, string> {
	const properties = inlineRunProperties(node);
	return Object.keys(properties).length
		? { 'data-run-properties': JSON.stringify(properties) }
		: {};
}

/** Keeps the imported basis stable and changes each direct property independently for Yjs. */
export function updatedInlineRunAttributes(node: Node, format: string | null): Node['attrs'] {
	if (!node.type.spec.attrs?.[attributeName('bold')]) return { ...node.attrs, format };
	const base: Record<string, unknown> =
		typeof node.attrs.format === 'string' ? JSON.parse(node.attrs.format) : {};
	const desired: Record<string, unknown> = format ? JSON.parse(format) : {};
	const metadata = { ...base };
	for (const key of new Set([...Object.keys(base), ...Object.keys(desired)])) {
		if (directKeys.has(key)) continue;
		if (Object.hasOwn(desired, key)) metadata[key] = desired[key];
		else delete metadata[key];
	}
	const attrs: Record<string, unknown> = {
		...node.attrs,
		format: Object.keys(metadata).length ? JSON.stringify(metadata) : null,
	};
	for (const key of DIRECT_RUN_PROPERTY_KEYS)
		attrs[attributeName(key)] =
			JSON.stringify(base[key]) === JSON.stringify(desired[key])
				? null
				: JSON.stringify(desired[key] ?? null);
	return attrs;
}

export function wordInlinePropertyCodec(
	schema: Node['type']['schema'],
): 'word-yjs-v1' | 'word-yjs-v2' {
	return Object.values(schema.nodes).some((type) => type.spec.attrs?.[attributeName('bold')])
		? 'word-yjs-v2'
		: 'word-yjs-v1';
}
