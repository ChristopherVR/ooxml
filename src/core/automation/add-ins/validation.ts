import type { AddInReadOptions, AddInRequirement } from './model';

export const MANIFEST_BYTES = 256 * 1024;
export const PACKAGE_BYTES = 4 * 1024 * 1024;
export const MANIFEST_DEPTH = 32;

export type AddInErrorCode =
	| 'manifest'
	| 'package'
	| 'limit'
	| 'url'
	| 'host'
	| 'version'
	| 'extension';

export class AddInManifestError extends Error {
	constructor(
		readonly code: AddInErrorCode,
		readonly field?: string,
	) {
		super(field ? `${code}: ${field}` : code);
		this.name = 'AddInManifestError';
	}
}

export function invalid(field: string): never {
	throw new AddInManifestError('manifest', field);
}

export function text(value: unknown, field: string, limit = 4096): string {
	if (
		typeof value !== 'string' ||
		!value.trim() ||
		value.length > limit ||
		/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f\u202a-\u202e\u2066-\u2069]/.test(value)
	)
		invalid(field);
	return value.trim();
}

export function guid(value: unknown, field: string): string {
	const id = text(value, field, 36);
	if (!/^[\da-f]{8}-(?:[\da-f]{4}-){3}[\da-f]{12}$/i.test(id)) invalid(field);
	return id.toLowerCase();
}

export function identifier(value: unknown, field: string): string {
	const id = text(value, field, 128);
	if (!/^[a-z\d_.-]+$/i.test(id)) invalid(field);
	return id;
}

export function url(value: unknown, field: string): string {
	if (
		typeof value === 'string' &&
		value.length <= 2048 &&
		/^https:\/\/[^/?#\\\s@]+/i.test(value) &&
		!/[\u0000-\u0020\u007f\\]/.test(value)
	) {
		const authority = value.split(/[/?#]/)[2] ?? '';
		try {
			const target = new URL(value);
			if (
				!authority.includes('@') &&
				!target.username &&
				!target.password &&
				!target.hostname.includes('*')
			)
				return target.href;
		} catch {
			throw new AddInManifestError('url', field);
		}
	}
	throw new AddInManifestError('url', field);
}

export const origin = (source: string): string => new URL(source).origin;

export function declaredOrigin(value: unknown, field: string, bare = false): string {
	const source = text(value, field, 2048);
	const target = url(bare ? `https://${source}` : source, field);
	const normalized = origin(target);
	// Check the literal path too, before URL normalization can erase dot segments.
	if (
		!/^https:\/\/[^/?#]+\/?$/i.test(bare ? `https://${source}` : source) ||
		target !== `${normalized}/`
	)
		throw new AddInManifestError('url', field);
	return normalized;
}

export function paneUrl(value: unknown, field: string, allowed: ReadonlySet<string>): string {
	const target = url(value, field);
	if (!allowed.has(origin(target))) throw new AddInManifestError('url', field);
	return target;
}

export function version(
	value: unknown,
	field: string,
	pattern = /^\d{1,5}(?:\.\d{1,5}){1,3}$/,
): string {
	const result = text(value, field, 23);
	if (!pattern.test(result)) invalid(field);
	return result;
}

function compareVersion(a: string, b: string): number {
	const left = a.split('.').map(Number);
	const right = b.split('.').map(Number);
	for (let i = 0; i < Math.max(left.length, right.length); i++) {
		const delta = (left[i] ?? 0) - (right[i] ?? 0);
		if (delta) return delta;
	}
	return 0;
}

export function requirement(name: unknown, min: unknown, max?: unknown): AddInRequirement {
	const result = {
		kind: 'set' as const,
		name: identifier(name, 'requirement name'),
		minVersion: version(min === undefined ? '1.1' : min, 'minimum requirement version'),
	};
	if (max === undefined) return result;
	const maxVersion = version(max, 'maximum requirement version');
	if (compareVersion(result.minVersion, maxVersion) > 0) invalid('requirement version range');
	return { ...result, maxVersion };
}

export function missingRequirements(
	requirements: readonly AddInRequirement[],
	{ capabilities }: AddInReadOptions,
): AddInRequirement[] {
	return requirements.filter(({ kind, name, minVersion, maxVersion }) => {
		const key = name.toLowerCase();
		if (kind === 'method') return !capabilities?.methods?.has(key);
		const supported = capabilities?.sets?.get(key);
		return (
			!supported ||
			compareVersion(supported, minVersion) < 0 ||
			(maxVersion !== undefined && compareVersion(supported, maxVersion) > 0)
		);
	});
}

export function unique<T>(items: readonly T[], key: (item: T) => string, field: string): void {
	const seen = new Set<string>();
	for (const item of items) {
		const id = key(item).toLowerCase();
		if (seen.has(id)) invalid(field);
		seen.add(id);
	}
}

export function localeTag(value: unknown): string {
	const tag = text(value, 'locale', 64);
	if (!/^[a-z]{2,8}(?:-[a-z0-9]{1,8})*$/i.test(tag)) invalid('locale');
	return tag.toLowerCase();
}

export function localeCandidates(locale: string | undefined): string[] {
	const parts = (locale ?? '').toLowerCase().replaceAll('_', '-').split('-');
	const candidates: string[] = [];
	while (parts.length && parts[0]) {
		candidates.push(parts.join('-'));
		parts.pop();
	}
	return candidates;
}

export const PERMISSIONS = new Map([
	['Restricted', 'Document.Restricted.User'],
	['ReadDocument', 'Document.Read.User'],
	['ReadAllDocument', 'Document.ReadAll.User'],
	['WriteDocument', 'Document.Write.User'],
	['ReadWriteDocument', 'Document.ReadWrite.User'],
]);
