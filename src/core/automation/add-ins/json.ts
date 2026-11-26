import { MANIFEST_DEPTH, invalid } from './validation';

/** Check keys before JSON.parse collapses duplicates, including equivalent escaped keys. */
export function parseManifestJson(source: string): unknown {
	const token =
		/\s*("(?:[^"\\\u0000-\u001f]|\\(?:["\\/bfnrt]|u[\da-fA-F]{4}))*"|[{}\[\]:,]|true|false|null|-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?)/y;
	let cursor = 0;
	const take = (): string => {
		token.lastIndex = cursor;
		const match = token.exec(source);
		if (!match?.[1]) return invalid('JSON syntax');
		cursor = token.lastIndex;
		return match[1];
	};
	const value = (next: string, depth: number): void => {
		if (next !== '{' && next !== '[') {
			if (/^[}\]:,]$/.test(next)) invalid('JSON value');
			return;
		}
		if (depth > MANIFEST_DEPTH) invalid('JSON nesting limit');
		const end = next === '{' ? '}' : ']';
		const keys = next === '{' ? new Set<string>() : null;
		let item = take();
		if (item === end) return;
		for (;;) {
			if (keys) {
				if (!item.startsWith('"')) invalid('JSON key');
				const key: string = JSON.parse(item);
				if (keys.has(key)) invalid('duplicate JSON key');
				keys.add(key);
				if (take() !== ':') invalid('JSON key separator');
				item = take();
			}
			value(item, depth + 1);
			const separator = take();
			if (separator === end) return;
			if (separator !== ',') invalid('JSON separator');
			item = take();
		}
	};
	value(take(), 1);
	try {
		return JSON.parse(source);
	} catch {
		return invalid('JSON syntax');
	}
}

type JsonObject = Readonly<Record<string, unknown>>;

export function object(value: unknown, field: string): JsonObject {
	if (!value || typeof value !== 'object' || Array.isArray(value)) return invalid(field);
	return value as JsonObject;
}

export function optionalObject(value: unknown, field: string): JsonObject {
	return value === undefined ? {} : object(value, field);
}

export function array(value: unknown, field: string): readonly unknown[] {
	if (!Array.isArray(value)) return invalid(field);
	return value;
}

export function optionalArray(value: unknown, field: string): readonly unknown[] {
	return value === undefined ? [] : array(value, field);
}

export function records(value: unknown, field: string): JsonObject[] {
	return optionalArray(value, field).map((item) => object(item, field));
}
